import streamlit as st
import os
import json
import zipfile
import tempfile
import shutil
import pymupdf as fitz
import docx
import openpyxl
import pytesseract
from PIL import Image
import io
import subprocess

st.set_page_config(page_title="Тендерный парсер", layout="wide")
st.title("📄 Парсер документации тендера")
st.write("Загрузите ZIP-архив с документами, чтобы получить JSON для Qwen.")

# --- Функции извлечения (те же, что были) ---
def extract_pdf(file_path):
    doc = fitz.open(file_path)
    full_text = []
    for page_num in range(len(doc)):
        page = doc[page_num]
        text = page.get_text()
        if not text.strip():
            try:
                pix = page.get_pixmap(dpi=300)
                img = Image.open(io.BytesIO(pix.tobytes("png")))
                text = pytesseract.image_to_string(img, lang='rus+eng')
            except Exception as e:
                text = f"[Ошибка OCR: {e}]"
        full_text.append(f"--- Страница {page_num + 1} ---\n{text}")
    doc.close()
    return "\n".join(full_text)

def extract_docx(file_path):
    doc = docx.Document(file_path)
    parts = [p.text for p in doc.paragraphs if p.text.strip()]
    for table in doc.tables:
        for row in table.rows:
            row_text = " | ".join(cell.text.strip() for cell in row.cells)
            if row_text.strip(): parts.append(row_text)
    return "\n".join(parts)

def extract_xlsx(file_path):
    wb = openpyxl.load_workbook(file_path, data_only=True)
    parts = []
    for sheet_name in wb.sheetnames:
        sheet = wb[sheet_name]
        parts.append(f"=== Лист: {sheet_name} ===")
        for row in sheet.iter_rows(values_only=True):
            row_text = " | ".join(str(cell) if cell is not None else "" for cell in row)
            if row_text.strip(): parts.append(row_text)
    wb.close()
    return "\n".join(parts)

def extract_doc_legacy(file_path):
    with tempfile.TemporaryDirectory() as tmpdir:
        subprocess.run(['libreoffice', '--headless', '--convert-to', 'docx', '--outdir', tmpdir, file_path], check=True, capture_output=True, timeout=120)
        base = os.path.splitext(os.path.basename(file_path))[0]
        new_path = os.path.join(tmpdir, base + '.docx')
        if os.path.exists(new_path): return extract_docx(new_path)
        return "[Не удалось конвертировать .doc]"

def extract_xls_legacy(file_path):
    with tempfile.TemporaryDirectory() as tmpdir:
        subprocess.run(['libreoffice', '--headless', '--convert-to', 'xlsx', '--outdir', tmpdir, file_path], check=True, capture_output=True, timeout=120)
        base = os.path.splitext(os.path.basename(file_path))[0]
        new_path = os.path.join(tmpdir, base + '.xlsx')
        if os.path.exists(new_path): return extract_xlsx(new_path)
        return "[Не удалось конвертировать .xls]"

# --- Интерфейс ---
uploaded_file = st.file_uploader("Выберите ZIP-архив с документами", type="zip")

if uploaded_file is not None:
    if st.button("🚀 Обработать архив"):
        with st.spinner("Распаковка и извлечение текста... Это может занять пару минут."):
            temp_dir = tempfile.mkdtemp()
            zip_path = os.path.join(temp_dir, "upload.zip")
            
            # Сохраняем загруженный файл
            with open(zip_path, "wb") as f:
                f.write(uploaded_file.getbuffer())
            
            # Распаковываем
            extract_dir = os.path.join(temp_dir, "extracted")
            with zipfile.ZipFile(zip_path, 'r') as zip_ref:
                zip_ref.extractall(extract_dir)
            
            # Собираем данные
            result = {"files": [], "total_files": 0, "errors": []}
            
            for root, dirs, files in os.walk(extract_dir):
                for filename in sorted(files):
                    file_path = os.path.join(root, filename)
                    rel_path = os.path.relpath(file_path, extract_dir)
                    ext = filename.lower().rsplit('.', 1)[-1] if '.' in filename else ''
                    file_info = {"filename": filename, "path": rel_path, "ext": ext, "text": ""}
                    
                    try:
                        if ext == 'pdf': file_info["text"] = extract_pdf(file_path)
                        elif ext == 'docx': file_info["text"] = extract_docx(file_path)
                        elif ext == 'xlsx': file_info["text"] = extract_xlsx(file_path)
                        elif ext == 'doc': file_info["text"] = extract_doc_legacy(file_path)
                        elif ext == 'xls': file_info["text"] = extract_xls_legacy(file_path)
                        elif ext in ('txt', 'csv'):
                            with open(file_path, 'r', encoding='utf-8', errors='ignore') as f: file_info["text"] = f.read()
                        else: file_info["text"] = f"[Формат .{ext} не поддерживается]"
                    except Exception as e:
                        file_info["text"] = f"[Ошибка: {e}]"
                        result["errors"].append({"file": rel_path, "error": str(e)})
                    
                    result["files"].append(file_info)
            
            result["total_files"] = len(result["files"])
            json_str = json.dumps(result, ensure_ascii=False, indent=2)
            
            # Очистка временных файлов
            shutil.rmtree(temp_dir)
            
            st.success(f"✅ Обработано файлов: {result['total_files']} (ошибок: {len(result['errors'])})")
            st.download_button(
                label="📥 Скачать JSON",
                file_name="tender_content.json",
                mime="application/json",
                data=json_str
            )
