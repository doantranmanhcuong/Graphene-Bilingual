import os
import streamlit as st
import openpyxl
import docx
from docx.shared import RGBColor, Pt
import time
import tempfile
import re
import json
from docx.oxml.ns import qn
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

class CloudDripEngine:
    def __init__(self, target_lang):
        self.target_lang = target_lang
        self.groq_keys = GROQ_API_KEYS
        self.current_key_idx = 0
        
    def _clean_json(self, text):
        text = text.strip()
        if text.startswith("```json"): text = text[7:]
        elif text.startswith("```"): text = text[3:]
        if text.endswith("```"): text = text[:-3]
        return text.strip()
        
    def _call_groq(self, prompt, status_text=None):
        if not self.groq_keys:
            raise ValueError("Hệ thống không có bất kỳ GROQ_API_KEY nào được cấu hình!")
            
        attempts = len(self.groq_keys)
        for i in range(attempts):
            api_key = self.groq_keys[self.current_key_idx]
            url = "https://api.groq.com/openai/v1/chat/completions"
            headers = {
                "Authorization": f"Bearer {api_key}",
                "Content-Type": "application/json"
            }
            data = {
                "model": "openai/gpt-oss-120b",
                "messages": [{"role": "user", "content": prompt}],
                "temperature": 0.01,
                "response_format": {"type": "json_object"}
            }
            
            try:
                response = requests.post(url, headers=headers, json=data, timeout=40)
                if response.status_code == 200:
                    return response.json()['choices'][0]['message']['content']
                elif response.status_code == 429: # Rate limit
                    if status_text:
                        status_text.text(f"🔄 Key Groq {self.current_key_idx + 1} bị Rate Limit. Đang xoay tua sang Key tiếp theo...")
                    self.current_key_idx = (self.current_key_idx + 1) % len(self.groq_keys)
                    continue
                else:
                    raise Exception(f"HTTP {response.status_code}: {response.text}")
            except Exception as e:
                if i == attempts - 1:
                    raise e
                self.current_key_idx = (self.current_key_idx + 1) % len(self.groq_keys)
                
        raise Exception("Tất cả các Key Groq đều thất bại hoặc bị Rate Limit!")

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
            
        draft_results = []
        # THUẬT TOÁN NHỎ GIỌT: Chỉ gửi 10 câu mỗi lần
        batch_size = 10 
        
        chunks = []
        current_chunk = []
        for text in unique_texts:
            current_chunk.append(text)
            if len(current_chunk) >= batch_size:
                chunks.append(current_chunk)
                current_chunk = []
        if current_chunk:
            chunks.append(current_chunk)
            
        total_chunks = len(chunks)
        for chunk_idx, chunk in enumerate(chunks):
            if progress_bar and status_text:
                progress_val = int(((chunk_idx) / total_chunks) * 100)
                progress_bar.progress(progress_val)
                status_text.text(f"🚀 Đang dịch cụm {chunk_idx+1}/{total_chunks}...")
            
            json_payload = {str(i): text for i, text in enumerate(chunk)}
            prompt = f"""You are a strict bilingual translation system. Translate the values of the following JSON object into {self.target_lang}.
CRITICAL RULES:
1. Translate accurately with NO additions, NO omissions.
2. Keep numbers, product codes, abbreviations EXACTLY as the original.
3. You MUST return ONLY a valid JSON object.
4. The output keys MUST match the input keys exactly.

Input JSON:
{json.dumps(json_payload, ensure_ascii=False)}"""

            success = False
            for attempt in range(2):
                try:
                    res_text = self._call_groq(prompt, status_text)
                    translated_dict = json.loads(res_text)
                    for i in range(len(chunk)):
                        draft_results.append(translated_dict.get(str(i), chunk[i]))
                    success = True
                    break
                except Exception as e:
                    if status_text:
                        status_text.text(f"⚠️ Groq nghẽn cụm {chunk_idx+1}. Thử lại (Lần {attempt+1}/2)...")
                    time.sleep(5)
                    
            if not success:
                # Nếu xui quá xui thử 2 lần đều tạch, giữ nguyên bản gốc
                draft_results.extend(chunk)
                
            # Nghỉ 4.5 giây giữa mỗi yêu cầu để xả Rate Limit của cả 2 hệ thống
            time.sleep(4.5) 

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
                    new_text = str(cell.value) + "\n" + trans
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
                    cell.value = str(cell.value) + " / " + trans
            
        filename = self._get_output_filename(file_path)
        output_path = os.path.join(TEMP_DIR, filename)
        wb.save(output_path)
        return output_path

    def process_word(self, file_path):
        import copy
        doc = docx.Document(file_path)
        texts = []
        paragraphs = []
        
        for p_element in doc.element.body.iter(qn('w:p')):
            p = Paragraph(p_element, doc)
            if p.text.strip():
                texts.append(p.text.strip())
                paragraphs.append(p)
                            
        translations = self.engine.translate_batch(texts, self.progress_bar, self.status_text)
            
        for p, trans in zip(paragraphs, translations):
            if trans.lower() != p.text.strip().lower():
                # Nhân bản paragraph để giữ nguyên TẤT CẢ định dạng (Justify, Bullet, Indent)
                new_p = copy.deepcopy(p._p)
                p._p.addnext(new_p)
                new_para = Paragraph(new_p, p._parent)
                new_para.clear() # Xóa chữ cũ
                
                # Ép đoạn văn gốc (Tiếng Việt) mất khoảng cách lề dưới (Space After)
                # Giúp đoạn Tiếng Anh dính sát ngay bên dưới, tạo thành 1 cụm song ngữ hoàn hảo
                try:
                    p.paragraph_format.space_after = Pt(0)
                except:
                    pass
                
                # Chèn bản dịch vào đoạn mới
                new_run = new_para.add_run(trans)
                new_run.italic = True
                new_run.font.color.rgb = RGBColor(0, 51, 102) 
                
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
                new_run = paragraph.add_run()
                new_run.text = "\n" + trans
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

st.set_page_config(page_title="Graphene Bilingual", page_icon="⚡", layout="centered")

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

with st.container():
    st.markdown("### ⚙️ Cấu hình hệ thống")
    col1, col2 = st.columns([1, 1])
    with col1:
        target_lang = st.selectbox("🌍 Ngôn ngữ đích:", ["English", "Vietnamese", "Japanese", "Korean", "Chinese (Simplified)", "ThaiLan"])
    with col2:
        st.info("💡 Hỗ trợ giữ nguyên định dạng: **Word, Excel.")

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
