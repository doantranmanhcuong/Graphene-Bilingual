import os
import streamlit as st

st.set_page_config(page_title="Graphene Bilingual", page_icon="⚡", layout="centered")
import openpyxl
import docx
from docx.shared import RGBColor, Pt
import time
import tempfile
import re
import json
from docx.oxml.ns import qn, nsdecls
from docx.oxml import parse_xml
from docx.text.paragraph import Paragraph
import requests

# Hỗ trợ xoay tua nhiều Key Groq để chống giới hạn Rate Limit
GROQ_API_KEYS = []

try:
    # 1. Thử lấy danh sách key từ st.secrets (Streamlit Cloud)
    keys_cloud = st.secrets.get("GROQ_API_KEYS", [])
    if isinstance(keys_cloud, list):
        GROQ_API_KEYS.extend(keys_cloud)
    elif keys_cloud:
        GROQ_API_KEYS.append(keys_cloud)
    key_cloud_single = st.secrets.get("GROQ_API_KEY", "")
    if key_cloud_single:
        GROQ_API_KEYS.append(key_cloud_single)
except:
    pass

try:
    from dotenv import load_dotenv
    load_dotenv()
    # 2. Lấy từ biến môi trường (Local)
    env_keys = os.getenv("GROQ_API_KEYS", "")
    if env_keys:
        GROQ_API_KEYS.extend([k.strip() for k in env_keys.split(',') if k.strip()])
    env_single = os.getenv("GROQ_API_KEY", "")
    if env_single:
        GROQ_API_KEYS.append(env_single)
except:
    pass

# 3. Bổ sung Key xoay tua của user (Đã chuyển sang bảo mật ở file .env)
# API Keys được nạp an toàn từ .env hoặc st.secrets. Tuyệt đối không hardcode trong mã nguồn.

# Lọc bỏ trùng lặp và rỗng
GROQ_API_KEYS = list(set([k for k in GROQ_API_KEYS if k]))

TEMP_DIR = os.path.join(tempfile.gettempdir(), "bilingual_cloud")
os.makedirs(TEMP_DIR, exist_ok=True)

import sqlite3
import hashlib
from concurrent.futures import ThreadPoolExecutor, as_completed
import queue
import threading
import re

class RateLimitException(Exception):
    def __init__(self, message, wait_time):
        self.message = message
        self.wait_time = wait_time
        super().__init__(self.message)

def get_db_connection():
    conn = sqlite3.connect(os.path.join(TEMP_DIR, 'translation_cache.db'), timeout=30.0)
    c = conn.cursor()
    c.execute('''CREATE TABLE IF NOT EXISTS cache
                 (hash_key TEXT PRIMARY KEY, target_lang TEXT, original TEXT, translated TEXT)''')
    # Tự động dọn dẹp các bản ghi lỗi từng bị lưu nhầm câu gốc vào cache
    c.execute("DELETE FROM cache WHERE TRIM(original) = TRIM(translated)")
    conn.commit()
    return conn

def get_cached_translation(conn, original, target_lang):
    hash_key = hashlib.md5((original + target_lang).encode('utf-8')).hexdigest()
    c = conn.cursor()
    c.execute("SELECT translated FROM cache WHERE hash_key=?", (hash_key,))
    row = c.fetchone()
    if row and row[0]:
        val = row[0]
        # Không chấp nhận bản dịch trùng y hệt câu gốc (tránh rác cache)
        if isinstance(val, str):
            if val.strip().lower() == original.strip().lower():
                return None
            val = val.replace('\\t', '\t').replace('\\r', '\r').replace('\\n', '\n')
        return val
    return None

def save_cached_translation(conn, original, target_lang, translated):
    if not translated or not original:
        return
    # Tuyệt đối không lưu vào cache nếu câu dịch giống hệt câu gốc (do API bị lỗi hoặc fallback)
    if translated.strip().lower() == original.strip().lower():
        return
    hash_key = hashlib.md5((original + target_lang).encode('utf-8')).hexdigest()
    try:
        c = conn.cursor()
        c.execute("INSERT OR REPLACE INTO cache (hash_key, target_lang, original, translated) VALUES (?, ?, ?, ?)",
                  (hash_key, target_lang, original, translated))
        conn.commit()
    except Exception:
        pass

class CloudDripEngine:
    def __init__(self, target_lang):
        self.target_lang = target_lang
        if not GROQ_API_KEYS:
            raise ValueError("Hệ thống không có bất kỳ GROQ_API_KEY nào được cấu hình!")
        self.keys = list(GROQ_API_KEYS)
        self.key_lock = threading.Lock()
        self.key_cooldowns = {k: 0.0 for k in self.keys}
        self.key_index = 0
        self.status_notes = ""

    def get_available_key(self):
        """Lấy một key đang sẵn sàng (không bị cooldown). Nếu tất cả đều cooldown, trả về (None, min_wait)."""
        with self.key_lock:
            now = time.time()
            available = [k for k in self.keys if self.key_cooldowns[k] <= now]
            if available:
                self.key_index = (self.key_index + 1) % len(available)
                chosen_key = available[self.key_index]
                if not any(self.key_cooldowns[k] > now for k in self.keys):
                    self.status_notes = ""
                return chosen_key, 0
            
            min_wait = max(1.0, min(self.key_cooldowns.values()) - now)
            return None, min_wait

    def mark_key_cooldown(self, api_key, wait_time):
        """Đưa key vào danh sách tạm nghỉ (cooldown) để nhường quyền cho các key khác."""
        with self.key_lock:
            now = time.time()
            self.key_cooldowns[api_key] = now + wait_time
            key_preview = api_key[:7] + "..." if len(api_key) > 10 else "Key"
            self.status_notes = f"Key {key_preview} tạm nghỉ ({int(wait_time)}s), chuyển key khác"

    def _clean_json(self, text):
        text = text.strip()
        if text.startswith("```json"): text = text[7:]
        elif text.startswith("```"): text = text[3:]
        if text.endswith("```"): text = text[:-3]
        return text.strip()
        
    def _call_groq_single(self, prompt, api_key):
        url = "https://api.groq.com/openai/v1/chat/completions"
        headers = {
            "Authorization": f"Bearer {api_key}",
            "Content-Type": "application/json"
        }
        data = {
            "model": "openai/gpt-oss-120b",
            "messages": [{"role": "user", "content": prompt}],
            "temperature": 0.01,
            "max_tokens": 4096,
            "response_format": {"type": "json_object"}
        }
        response = requests.post(url, headers=headers, json=data, timeout=60)
        if response.status_code == 200:
            return response.json()['choices'][0]['message']['content']
        elif response.status_code == 429:
            try:
                err_msg = response.json().get("error", {}).get("message", "")
                match = re.search(r'try again in (?:(\d+)h)?(?:(\d+)m)?(?:([\d.]+)s)?', err_msg)
                wait_time = 15
                if match:
                    h, m, s = match.groups()
                    total = 0
                    if h: total += int(h) * 3600
                    if m: total += int(m) * 60
                    if s: total += float(s)
                    wait_time = total + 3 # Đệm 3s an toàn
                raise RateLimitException(err_msg, wait_time)
            except RateLimitException as re_err:
                raise re_err
            except Exception:
                raise RateLimitException(response.text, 15)
        else:
            raise Exception(f"HTTP {response.status_code}: {response.text}")

    def _process_chunk_worker(self, chunk_idx, chunk):
        if not chunk:
            return []

        PROMPT_LANG_MAP = {
            "ThaiLan": "Thai",
            "Chinese (Simplified)": "Simplified Chinese",
        }
        lang_name = PROMPT_LANG_MAP.get(self.target_lang, self.target_lang)

        json_payload = {str(i): text for i, text in enumerate(chunk)}
        prompt = f"""You are an expert professional translator specializing in legal, corporate, and formal technical documents (including SDS and contracts). Translate the values of the following JSON object into {lang_name}.

CRITICAL RULES:
1. FORMAL & NATURAL TONE: The translation must be highly professional, fluent, and suitable for official/technical contexts. Avoid clunky, word-for-word, or robotic translations. Use contextually appropriate terminology according to standards (e.g., GHS / ABNT NBR 14725 for SDS).
2. TRANSLITERATE PROPER NAMES: When translating into non-Latin scripts (such as Thai, Chinese, Japanese, Russian, Korean), transliterate all personal names, organization/company names, titles, and place names phonetically into the target script (e.g., Thai script). NEVER leave names untranslated in Latin letters. When the target language uses Latin script (e.g. English, Portuguese), keep standard romanized names.
3. STRICT FIDELITY: Absolutely NO hallucinations, additions, or omissions. Translate the exact meaning of the original text without making anything up.
4. PRESERVE FORMATTING: Keep all numbers, product codes, abbreviations, and special characters EXACTLY as they appear in the original.
5. NO ESCAPED SEQUENCES: Do NOT output visible escaped sequences like '\\t' or '\\n' as literal text. Maintain natural spacing and column alignment.
6. JSON OUTPUT ONLY: You MUST return ONLY a valid JSON object. The output keys MUST match the input keys exactly.

Input JSON:
{json.dumps(json_payload, ensure_ascii=False)}"""

        max_attempts = 4
        for attempt in range(max_attempts):
            api_key, wait_time = self.get_available_key()
            if not api_key:
                if wait_time > 60:
                    raise RateLimitException(f"Tất cả {len(self.keys)} Key Groq đã tạm hết hạn mức (cần nghỉ {int(wait_time)}s). Tiến độ đã lưu an toàn vào Cache.", wait_time)
                time.sleep(min(wait_time, 5))
                continue

            try:
                res_text = self._call_groq_single(prompt, api_key)
                cleaned = self._clean_json(res_text)
                translated_dict = json.loads(cleaned)
                
                results = []
                for i in range(len(chunk)):
                    val = translated_dict.get(str(i), chunk[i])
                    if isinstance(val, str):
                        val = val.replace('\\t', '\t').replace('\\r', '\r').replace('\\n', '\n')
                    results.append(val)
                time.sleep(0.3)
                return results

            except RateLimitException as e:
                # Key này bị giới hạn -> đưa vào cooldown và ngay lập tức lấy Key khác ở lượt kế
                self.mark_key_cooldown(api_key, e.wait_time)
                continue
            except Exception:
                if attempt < max_attempts - 1:
                    time.sleep(1.5)
                else:
                    # Nếu chunk nhiều câu bị lỗi (quá tải token hoặc parse fail), chia đôi để dịch từng nửa
                    if len(chunk) > 1:
                        mid = len(chunk) // 2
                        left_res = self._process_chunk_worker(f"{chunk_idx}_a", chunk[:mid])
                        right_res = self._process_chunk_worker(f"{chunk_idx}_b", chunk[mid:])
                        return left_res + right_res
                    else:
                        return chunk

        if len(chunk) > 1:
            mid = len(chunk) // 2
            left_res = self._process_chunk_worker(f"{chunk_idx}_a", chunk[:mid])
            right_res = self._process_chunk_worker(f"{chunk_idx}_b", chunk[mid:])
            return left_res + right_res

        return chunk

    def translate_batch(self, texts, progress_bar=None, status_text=None):
        if not texts:
            return []
            
        processed_results = [None] * len(texts)
        to_translate = []
        to_translate_indices = []
        
        for i, text in enumerate(texts):
            stripped = text.strip()
            if not stripped or bool(re.match(r'^[\d.,\s\-\+/%:]+$', stripped)):
                processed_results[i] = text
            else:
                to_translate.append(text)
                to_translate_indices.append(i)
                
        unique_texts = []
        text_to_unique_idx = []
        for text in to_translate:
            if text not in unique_texts:
                text_to_unique_idx.append(len(unique_texts))
                unique_texts.append(text)
            else:
                text_to_unique_idx.append(unique_texts.index(text))
                
        if not unique_texts:
            return processed_results
            
        conn = get_db_connection()
        cached_results = {}
        missing_texts = []
        
        for idx, text in enumerate(unique_texts):
            cached = get_cached_translation(conn, text, self.target_lang)
            if cached:
                cached_results[text] = cached
            else:
                missing_texts.append(text)
                
        if not missing_texts:
            if status_text:
                status_text.text("⚡ Phục hồi 100% từ bộ nhớ đệm (Cache)...")
            draft_results = [cached_results.get(text, text) for text in unique_texts]
            
            for i, original_idx in enumerate(to_translate_indices):
                unique_idx = text_to_unique_idx[i]
                processed_results[original_idx] = draft_results[unique_idx]
            return processed_results

        # Gom cụm thông minh theo cả số câu VÀ tổng độ dài ký tự
        # Đối với hợp đồng/văn bản dài: các đoạn văn dài hàng trăm chữ nếu gom 20 câu sẽ vượt ngưỡng token
        # Giới hạn an toàn: tối đa 8 đoạn HOẶC tối đa 1800 ký tự mỗi cụm
        max_items_per_chunk = 8
        max_chars_per_chunk = 1800
        
        chunks = []
        current_chunk = []
        current_chars = 0
        
        for text in missing_texts:
            text_len = len(text)
            if current_chunk and (len(current_chunk) >= max_items_per_chunk or (current_chars + text_len) > max_chars_per_chunk):
                chunks.append(current_chunk)
                current_chunk = []
                current_chars = 0
            current_chunk.append(text)
            current_chars += text_len
            
        if current_chunk:
            chunks.append(current_chunk)
            
        total_chunks = len(chunks)
        max_workers = min(len(self.keys), 4)
        completed_chunks = 0
        
        if status_text:
            status_text.text(f"🚀 Kích hoạt Đa Luồng ({max_workers} luồng). Đang xử lý {total_chunks} cụm...")

        # Đính kèm ScriptRunContext để thread pool hoạt động chuẩn trong Streamlit
        try:
            from streamlit.runtime.scriptrunner import add_script_run_ctx, get_script_run_ctx
            ctx = get_script_run_ctx()
        except Exception:
            ctx = None

        def worker_task(chunk_idx, chunk):
            if ctx:
                try:
                    add_script_run_ctx(ctx=ctx)
                except Exception:
                    pass
            return self._process_chunk_worker(chunk_idx, chunk)
            
        futures = {}
        with ThreadPoolExecutor(max_workers=max_workers) as executor:
            for chunk_idx, chunk in enumerate(chunks):
                future = executor.submit(worker_task, chunk_idx, chunk)
                futures[future] = (chunk_idx, chunk)
                
            for future in as_completed(futures):
                chunk_idx, chunk = futures[future]
                results = future.result()
                
                for orig, trans in zip(chunk, results):
                    if trans and trans.strip().lower() != orig.strip().lower():
                        save_cached_translation(conn, orig, self.target_lang, trans)
                        cached_results[orig] = trans
                    else:
                        cached_results[orig] = orig
                    
                completed_chunks += 1
                if progress_bar and status_text:
                    progress_val = int((completed_chunks / total_chunks) * 100)
                    progress_bar.progress(progress_val)
                    note = f" ({self.status_notes})" if self.status_notes else ""
                    status_text.text(f"🚀 Tiến độ: {completed_chunks}/{total_chunks} cụm ({progress_val}%){note}")

        draft_results = []
        for text in unique_texts:
            draft_results.append(cached_results.get(text, text))

        for i, original_idx in enumerate(to_translate_indices):
            unique_idx = text_to_unique_idx[i]
            if unique_idx < len(draft_results):
                processed_results[original_idx] = draft_results[unique_idx]
            else:
                processed_results[original_idx] = unique_texts[unique_idx]
                
        return processed_results


class DocumentProcessor:
    def __init__(self, target_lang, progress_bar=None, status_text=None):
        self.engine = CloudDripEngine(target_lang)
        self.progress_bar = progress_bar
        self.status_text = status_text

    def _get_output_filename(self, file_path, target_ext=None):
        base_name = os.path.basename(file_path)
        name, ext = os.path.splitext(base_name)
        if target_ext is None:
            target_ext = ext
        return f"{name}_Bilingual{target_ext}"

    def _calculate_exact_height(self, sheet, cell, text):
        from openpyxl.utils import get_column_letter
        import math
        
        # 1. Đo chiều rộng thực tế của ô (Bao gồm cả tính toán gộp ô - Merged Cells)
        effective_width = 10.0 
        is_merged = False
        
        for merged_range in sheet.merged_cells.ranges:
            if cell.coordinate in merged_range:
                is_merged = True
                total_width = 0
                for col_idx in range(merged_range.min_col, merged_range.max_col + 1):
                    col_letter = get_column_letter(col_idx)
                    col_dim = sheet.column_dimensions.get(col_letter)
                    total_width += col_dim.width if (col_dim and col_dim.width) else 10.0
                effective_width = total_width
                break
                
        if not is_merged:
            col_letter = cell.column_letter
            col_dim = sheet.column_dimensions.get(col_letter)
            effective_width = col_dim.width if (col_dim and col_dim.width) else 10.0
            
        # 2. Ước tính số ký tự tối đa chứa được trên 1 dòng vật lý
        # Hệ số an toàn: Giảm xuống 0.85 (thay vì 1.1) để trừ hao cho các chữ in hoa và chữ có dấu Tiếng Việt
        max_chars_per_line = max(int(effective_width * 0.85), 5)
        
        # 3. Chạy vòng lặp mô phỏng thuật toán Wrap Text của MS Excel
        total_lines = 0
        for line in text.split('\n'):
            if len(line) == 0:
                total_lines += 1
            else:
                total_lines += math.ceil(len(line) / max_chars_per_line)
                
        # 4. Tính ra chiều cao chính xác (Cộng thêm 0.6 dòng làm vùng đệm padding để chắc chắn không bao giờ bị lẹm nét chữ dưới đáy)
        return (total_lines + 0.6) * 15

    def process_excel(self, file_path):
        from openpyxl.styles import Alignment
        import copy
        
        wb = openpyxl.load_workbook(file_path)
        texts = []
        cells = []
        for sheet in wb.worksheets:
            for row in sheet.iter_rows():
                for cell in row:
                    if cell.value and isinstance(cell.value, str) and cell.value.strip():
                        texts.append(cell.value.strip())
                        cells.append(cell)
                        
        translations = self.engine.translate_batch(texts, self.progress_bar, self.status_text)
        
        for cell, trans in zip(cells, translations):
            if trans.lower() != str(cell.value).strip().lower():
                clean_trans = trans.replace('\\t', '    ').replace('\\r', '').replace('\\n', '\n')
                is_explicitly_wrapped = cell.alignment and cell.alignment.wrap_text
                
                # Kiểm tra xem ô có thuộc Merged Cells không
                is_merged = False
                for merged_range in sheet.merged_cells.ranges:
                    if cell.coordinate in merged_range:
                        is_merged = True
                        break
                        
                # Kiểm tra xem ô bên phải có trống không (nếu trống -> có thể tràn ngang làm tiêu đề)
                next_cell = sheet.cell(row=cell.row, column=cell.column + 1)
                can_overflow = (next_cell.value is None or str(next_cell.value).strip() == "")
                
                # Quyết định Wrap Text thông minh dựa vào ngữ cảnh
                should_wrap = is_explicitly_wrapped or is_merged or (not can_overflow)
                
                if should_wrap:
                    # Chèn bản dịch xuống dòng
                    new_text = str(cell.value) + "\n" + clean_trans
                    cell.value = new_text
                    
                    if cell.alignment:
                        new_align = copy.copy(cell.alignment)
                        new_align.wrap_text = True
                        cell.alignment = new_align
                    else:
                        cell.alignment = Alignment(wrap_text=True)
                    
                    # Gọi Thuật toán AI mô phỏng Rendering của Excel để ép chiều cao chuẩn xác
                    calculated_height = self._calculate_exact_height(sheet, cell, new_text)
                    current_height = sheet.row_dimensions[cell.row].height
                    
                    if current_height is None or calculated_height > current_height:
                        sheet.row_dimensions[cell.row].height = calculated_height
                else:
                    # Các ô tiêu đề không gộp ô và có không gian trống để tràn ngang -> Dùng dấu " / "
                    cell.value = str(cell.value) + " / " + clean_trans
            
        filename = self._get_output_filename(file_path)
        output_path = os.path.join(TEMP_DIR, filename)
        wb.save(output_path)
        return output_path

    def _detect_doc_default_font(self, doc):
        try:
            rFonts = doc.styles.element.xpath('.//w:docDefaults//w:rFonts')
            if rFonts:
                val = rFonts[0].get(qn('w:ascii')) or rFonts[0].get(qn('w:hAnsi')) or rFonts[0].get(qn('w:cs'))
                if val:
                    return val
        except Exception:
            pass
        try:
            if doc.styles['Normal'].font.name:
                return doc.styles['Normal'].font.name
        except Exception:
            pass
        return "Times New Roman"

    def _set_run_font(self, run, font_name):
        if font_name:
            run.font.name = font_name
            rPr = run._r.get_or_add_rPr()
            rFonts = rPr.get_or_add_rFonts()
            rFonts.set(qn('w:ascii'), font_name)
            rFonts.set(qn('w:hAnsi'), font_name)
            rFonts.set(qn('w:cs'), font_name)

    def process_word(self, file_path):
        import copy
        doc = docx.Document(file_path)
        default_doc_font = self._detect_doc_default_font(doc)
        
        texts = []
        paragraphs = []
        
        # 1. Thu thập toàn bộ paragraph trong phần thân văn bản (bao gồm tất cả bảng biểu)
        for p_element in doc.element.body.iter(qn('w:p')):
            p = Paragraph(p_element, doc)
            if p.text.strip():
                texts.append(p.text.strip())
                paragraphs.append(p)
                
        # 2. Thu thập thêm paragraph trong Header và Footer (nếu có bảng biểu hoặc văn bản)
        for section in doc.sections:
            for hf in [section.header, section.footer]:
                if hf and not hf.is_linked_to_previous:
                    for p_element in hf._element.iter(qn('w:p')):
                        p = Paragraph(p_element, doc)
                        if p.text.strip() and p not in paragraphs:
                            texts.append(p.text.strip())
                            paragraphs.append(p)
                            
        translations = self.engine.translate_batch(texts, self.progress_bar, self.status_text)
            
        for p, trans in zip(paragraphs, translations):
            if trans.lower() != p.text.strip().lower():
                clean_trans = trans.replace('\\t', '\t').replace('\\r', '\r').replace('\\n', '\n')
                
                # Trích xuất kiểu dáng font, cỡ chữ, độ đậm và màu sắc từ đoạn gốc
                orig_font_name = None
                orig_font_size = None
                orig_is_bold = False
                is_white_or_light = False
                
                for r in p.runs:
                    if r.text.strip():
                        if r.font.name and not orig_font_name:
                            orig_font_name = r.font.name
                        if r.font.size and not orig_font_size:
                            orig_font_size = r.font.size
                        if r.bold:
                            orig_is_bold = True
                        if r.font.color and r.font.color.rgb:
                            rgb = r.font.color.rgb
                            if rgb == RGBColor(255, 255, 255) or (rgb[0] > 200 and rgb[1] > 200 and rgb[2] > 200):
                                is_white_or_light = True
                                
                # Kiểm tra màu nền tối của paragraph (shading)
                shd = p._p.xpath('.//w:pPr/w:shd')
                if shd:
                    shd_val = shd[0].get(qn('w:fill'))
                    if shd_val and shd_val.lower() not in ['auto', 'none', 'ffffff']:
                        is_white_or_light = True
                
                is_in_table = bool(p._p.xpath('ancestor::w:tc'))
                if is_in_table:
                    tc_shd = p._p.xpath('ancestor::w:tc[1]//w:tcPr/w:shd')
                    if tc_shd:
                        tc_shd_val = tc_shd[0].get(qn('w:fill'))
                        if tc_shd_val and tc_shd_val.lower() not in ['auto', 'none', 'ffffff']:
                            is_white_or_light = True

                font_to_apply = orig_font_name or default_doc_font

                # Nhân bản paragraph để giữ nguyên TẤT CẢ định dạng khối (Justify, Indent, Tabs...)
                new_p = copy.deepcopy(p._p)
                p._p.addnext(new_p)
                new_para = Paragraph(new_p, p._parent)
                new_para.clear() # Xóa nội dung run cũ
                
                # --- KHÓA CHẶT ĐOẠN GỐC VÀ BẢN DỊCH TRÊN CÙNG MỘT TRANG (CHỐNG TRÀN TÁCH TRANG) ---
                # Đảm bảo Word không bao giờ ngắt trang chen ngang giữa câu gốc và câu dịch
                try:
                    p.paragraph_format.keep_with_next = True
                except Exception:
                    pass
                    
                # Xóa keepNext trên đoạn dịch để Word tự do ngắt trang sau bản dịch khi hết trang
                for kn in new_p.xpath('.//w:pPr/w:keepNext'):
                    kn.getparent().remove(kn)

                # --- TRIỆT TIÊU ĐÁNH SỐ TỰ ĐỘNG TRÊN ĐOẠN DỊCH ---
                # Tránh làm đẩy số thứ tự của các điều khoản/khoản mục (ví dụ Điều 76 -> Điều 77, 1. -> 2.)
                # 1. Xóa bỏ numPr sao chép từ đoạn gốc (nếu có)
                for np in new_p.xpath('.//w:pPr/w:numPr'):
                    np.getparent().remove(np)
                    
                # 2. Thêm numPr với numId=0 để tắt đánh số tự động kế thừa từ Style (như Heading 2, List Paragraph)
                new_numPr = parse_xml(f'<w:numPr {nsdecls("w")}><w:numId {nsdecls("w")} w:val="0"/></w:numPr>')
                pPr = new_p.xpath('.//w:pPr')
                if pPr:
                    pPr[0].append(new_numPr)
                    
                # 3. Điều chỉnh căn lề thụt đầu dòng (Indentation):
                # Nếu đoạn gốc có lề treo (hanging indent, ví dụ số mục ở 0, chữ ở 567 twips),
                # xóa hanging để dòng đầu của bản dịch thẳng hàng hoàn hảo với chữ của mục gốc
                ind_elems = new_p.xpath('.//w:pPr/w:ind')
                if ind_elems:
                    ind = ind_elems[0]
                    if qn('w:hanging') in ind.attrib:
                        del ind.attrib[qn('w:hanging')]
                
                # Điều chỉnh khoảng cách dòng: Đoạn gốc và bản dịch gắn kết chặt chẽ
                try:
                    orig_after = p.paragraph_format.space_after
                    p.paragraph_format.space_after = Pt(1)
                    if is_in_table:
                        new_para.paragraph_format.space_after = Pt(1)
                        new_para.paragraph_format.space_before = Pt(0)
                    else:
                        if orig_after is not None and orig_after > Pt(0):
                            new_para.paragraph_format.space_after = orig_after
                        else:
                            new_para.paragraph_format.space_after = Pt(4)
                except Exception:
                    pass
                
                # Thêm văn bản dịch vào run mới
                new_run = new_para.add_run(clean_trans)
                new_run.italic = True
                if orig_is_bold:
                    new_run.bold = True
                    
                # Điều chỉnh cỡ chữ: Trong bảng biểu thu nhỏ vừa phải để tránh xô lệch độ rộng cột
                if is_in_table:
                    if orig_font_size and orig_font_size > Pt(9):
                        new_run.font.size = max(Pt(8), orig_font_size - Pt(1.5))
                    else:
                        new_run.font.size = Pt(8.5)
                else:
                    if orig_font_size:
                        new_run.font.size = orig_font_size
                    
                self._set_run_font(new_run, font_to_apply)
                
                # Thiết lập mã ngôn ngữ chuẩn (w:lang) để Word kích hoạt ngắt từ tự nhiên (ví dụ ngắt từ Tiếng Thái)
                LANG_CODE_MAP = {
                    "ThaiLan": "th-TH",
                    "Chinese (Simplified)": "zh-CN",
                    "Japanese": "ja-JP",
                    "Korean": "ko-KR",
                    "Portuguese": "pt-BR",
                    "English": "en-US",
                    "Vietnamese": "vi-VN",
                }
                lang_code = LANG_CODE_MAP.get(self.engine.target_lang, "en-US")
                rPr = new_run._r.get_or_add_rPr()
                lang_elem = rPr.find(qn('w:lang'))
                if lang_elem is None:
                    lang_elem = parse_xml(f'<w:lang {nsdecls("w")} w:val="{lang_code}" w:bidi="{lang_code}" w:eastAsia="{lang_code}"/>')
                    rPr.append(lang_elem)
                else:
                    lang_elem.set(qn('w:val'), lang_code)
                    lang_elem.set(qn('w:bidi'), lang_code)
                    lang_elem.set(qn('w:eastAsia'), lang_code)
                
                # Lựa chọn màu sắc thông minh:
                if is_white_or_light:
                    # Nền tối: dùng màu xanh sáng (#E0F2FE) để nổi bật rực rỡ, không bị chìm
                    new_run.font.color.rgb = RGBColor(224, 242, 254)
                else:
                    # Nền sáng: dùng màu xanh navy trang trọng, chuyên nghiệp
                    new_run.font.color.rgb = RGBColor(0, 51, 102)

        # 3. Khóa chống ngắt dòng giữa trang cho tất cả các hàng trong bảng biểu (w:cantSplit)
        for table in doc.tables:
            for row in table.rows:
                trPr = row._tr.get_or_add_trPr()
                if not trPr.xpath('./w:cantSplit'):
                    trPr.append(parse_xml(f'<w:cantSplit {nsdecls("w")}/>'))
                
        filename = self._get_output_filename(file_path)
        output_path = os.path.join(TEMP_DIR, filename)
        doc.save(output_path)
        return output_path

    def process_powerpoint(self, file_path):
        import pptx
        from pptx.dml.color import RGBColor as PptxRGBColor
        prs = pptx.Presentation(file_path)
        
        texts = []
        shapes_to_translate = []
        
        for slide in prs.slides:
            for shape in slide.shapes:
                if hasattr(shape, "text") and shape.has_text_frame:
                    for paragraph in shape.text_frame.paragraphs:
                        if paragraph.text.strip():
                            texts.append(paragraph.text.strip())
                            shapes_to_translate.append(paragraph)
                            
        translations = self.engine.translate_batch(texts, self.progress_bar, self.status_text)
        
        for paragraph, trans in zip(shapes_to_translate, translations):
            if trans.lower() != paragraph.text.strip().lower():
                clean_trans = trans.replace('\\t', '\t').replace('\\r', '\r').replace('\\n', '\n')
                new_run = paragraph.add_run()
                new_run.text = "\n" + clean_trans
                new_run.font.italic = True
                new_run.font.color.rgb = PptxRGBColor(0, 51, 102)
                
        filename = self._get_output_filename(file_path)
        output_path = os.path.join(TEMP_DIR, filename)
        prs.save(output_path)
        return output_path

    def process_txt(self, file_path):
        with open(file_path, "r", encoding="utf-8") as f:
            lines = f.readlines()
        
        texts = [line.strip() for line in lines if line.strip()]
        translations = self.engine.translate_batch(texts, self.progress_bar, self.status_text)
        
        output_lines = []
        trans_idx = 0
        for line in lines:
            output_lines.append(line)
            if line.strip():
                output_lines.append(translations[trans_idx] + "\n\n")
                trans_idx += 1
                
        filename = self._get_output_filename(file_path)
        output_path = os.path.join(TEMP_DIR, filename)
        with open(output_path, "w", encoding="utf-8") as f:
            f.writelines(output_lines)
        return output_path

st.markdown("""
<style>
    @import url('https://fonts.googleapis.com/css2?family=Inter:wght@400;500;600;800&display=swap');
    
    html, body, [class*="css"] {
        font-family: 'Inter', sans-serif;
    }
    
    .main-title {
        text-align: center;
        background: linear-gradient(135deg, #3b82f6 0%, #06b6d4 100%);
        -webkit-background-clip: text;
        -webkit-text-fill-color: transparent;
        font-weight: 800;
        font-size: 3rem;
        margin-bottom: 0px;
        padding-bottom: 0px;
    }
    
    .sub-title {
        text-align: center;
        opacity: 0.7;
        font-weight: 500;
        font-size: 1.1rem;
        margin-bottom: 30px;
        margin-top: 5px;
    }
    
    .stButton>button {
        background: linear-gradient(135deg, #3b82f6 0%, #2563eb 100%);
        color: white;
        font-weight: 600;
        border: none;
        border-radius: 8px;
        padding: 0.6rem 2rem;
        transition: all 0.3s ease;
        box-shadow: 0 4px 6px -1px rgba(59, 130, 246, 0.3);
    }
    .stButton>button:hover {
        transform: translateY(-2px);
        box-shadow: 0 10px 15px -3px rgba(59, 130, 246, 0.4);
        background: linear-gradient(135deg, #2563eb 0%, #1d4ed8 100%);
        border: none;
        color: white;
    }
    
    /* Upload box styling */
    .stFileUploader>div>div {
        border: 2px dashed #3b82f6;
        border-radius: 12px;
        padding: 20px;
        transition: all 0.3s;
    }
    .stFileUploader>div>div:hover {
        border-color: #2563eb;
    }
</style>
""", unsafe_allow_html=True)

st.markdown("<h1 class='main-title'>⚡ Graphene Bilingual</h1>", unsafe_allow_html=True)
st.markdown("<p class='sub-title'>Công cụ dịch thuật tài liệu đa định dạng siêu tốc</p>", unsafe_allow_html=True)

if not GROQ_API_KEYS:
    st.error("⚠️ Hệ thống chưa được cấu hình GROQ_API_KEY. Vui lòng thêm vào file .env hoặc tab Secrets của Streamlit Cloud!")

LANGUAGE_OPTIONS = {
    "English": "English (Tiếng Anh)",
    "Vietnamese": "Vietnamese (Tiếng Việt)",
    "Japanese": "Japanese (Tiếng Nhật)",
    "Korean": "Korean (Tiếng Hàn)",
    "Chinese (Simplified)": "Chinese (Tiếng Trung Giản Thể)",
    "ThaiLan": "Thai (Tiếng Thái)",
    "Portuguese": "Portuguese (Tiếng Bồ Đào Nha)"
}

with st.container():
    st.markdown("### ⚙️ Cấu hình hệ thống")
    col1, col2 = st.columns([1, 1])
    with col1:
        target_lang = st.selectbox(
            "🌍 Ngôn ngữ đích:",
            options=list(LANGUAGE_OPTIONS.keys()),
            format_func=lambda x: LANGUAGE_OPTIONS.get(x, x)
        )
    with col2:
        st.info("💡 Hỗ trợ giữ nguyên định dạng: **Word, Excel**.")

with st.container():
    st.markdown("### 📂 Tải tài liệu lên")
    uploaded_file = st.file_uploader("Kéo thả hoặc chọn file từ máy tính của bạn", type=["docx", "xlsx"], label_visibility="collapsed")

if uploaded_file is not None:
    st.markdown("<br>", unsafe_allow_html=True)
    col_btn1, col_btn2, col_btn3 = st.columns([1, 2, 1])
    with col_btn2:
        start_btn = st.button("🚀 KÍCH HOẠT DỊCH THUẬT", use_container_width=True)
        
    if start_btn:
        start_time = time.time()
        
        file_path = os.path.join(TEMP_DIR, uploaded_file.name)
        with open(file_path, "wb") as f:
            f.write(uploaded_file.getbuffer())
            
        ext = os.path.splitext(file_path)[1].lower()
        
        st.markdown("---")
        progress_bar = st.progress(0)
        status_text = st.empty()
        
        processor = DocumentProcessor(target_lang, progress_bar, status_text)
        
        try:
            output_path = None
            if ext == '.xlsx':
                output_path = processor.process_excel(file_path)
            elif ext == '.docx':
                output_path = processor.process_word(file_path)
            elif ext == '.pptx':
                output_path = processor.process_powerpoint(file_path)
            elif ext == '.txt':
                output_path = processor.process_txt(file_path)
                
            progress_bar.progress(100)
            status_text.success(f"✅ Hoàn thành bản dịch (Thời gian: {time.time() - start_time:.2f}s)")
            
            with open(output_path, "rb") as file:
                btn = st.download_button(
                    label="⬇️ TẢI FILE SONG NGỮ VỀ MÁY",
                    data=file,
                    file_name=os.path.basename(output_path),
                    mime="application/octet-stream",
                    use_container_width=True
                )
                
        except Exception as e:
            st.error(f"❌ Lỗi Hệ Thống: {str(e)}")
