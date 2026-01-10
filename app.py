import streamlit as st
import os
import re
import camelot
import fitz  # PyMuPDF
import pdfplumber
from docx import Document
import pandas as pd
import io

# --- 公共函數部分 ---

def extract_values_from_filename(filename):
    values = re.findall(r'\d+', filename)
    if len(values) >= 3:
        return values[:3]
    return None

def extract_table_value(pdf_path, page_num, row_num, col_num):
    try:
        tables = camelot.read_pdf(pdf_path, pages=str(page_num), flavor='stream')
        for table in tables:
            df = table.df
            try:
                value = df.iat[int(row_num), int(col_num)].replace(',', '').replace(' ', '')
                return value
            except IndexError:
                continue
    except Exception:
        pass
    return "N/A"

def extract_row_values(pdf_path, page_num, keyword):
    try:
        tables = camelot.read_pdf(pdf_path, pages=str(page_num), flavor='stream')
        for table in tables:
            df = table.df
            for i, row in df.iterrows():
                if keyword in row.to_string():
                    values = [val.replace(',', '') for val in re.findall(r"[\d,.]+", row.to_string())]
                    return values
    except Exception:
        pass
    return []

def add_thousand_separator(value):
    try:
        value = float(value)
        if value.is_integer():
            formatted_value = "{:,.0f}".format(value)
        else:
            formatted_value = "{:,.1f}".format(value)
        return formatted_value
    except ValueError:
        return value

def evaluate_expression(expression, values):
    for key, value in values.items():
        expression = expression.replace(f"{{{key}}}", str(value))
    try:
        result = eval(expression, {"__builtins__": None}, {})
        return add_thousand_separator(result)
    except Exception:
        return "N/A"

def replace_and_evaluate_in_run(run, values):
    full_text = run.text
    for key, value in values.items():
        placeholder = f"{{{key}}}"
        full_text = full_text.replace(placeholder, str(value) if value is not None else "N/A")

    expressions = re.findall(r'\{\{[^\}]+\}\}', full_text)
    for expr in expressions:
        expr_clean = expr.strip("{}")
        result = evaluate_expression(expr_clean, values)
        full_text = full_text.replace(expr, result)

    run.text = full_text

def replace_and_evaluate_in_paragraph(paragraph, values):
    for run in paragraph.runs:
        replace_and_evaluate_in_run(run, values)

def process_word_template(template_path, values, remove_text_start=None, remove_text_end=None):
    doc = Document(template_path)
    for paragraph in doc.paragraphs:
        replace_and_evaluate_in_paragraph(paragraph, values)
    for table in doc.tables:
        for row in table.rows:
            for cell in row.cells:
                for paragraph in cell.paragraphs:
                    replace_and_evaluate_in_paragraph(paragraph, values)
    if remove_text_start and remove_text_end:
        delete_specified_range(doc, remove_text_start, remove_text_end)
    bio = io.BytesIO()
    doc.save(bio)
    bio.seek(0)
    return bio

def delete_specified_range(doc, start_text, end_text):
    paragraphs = list(doc.paragraphs)
    start_idx = -1
    end_idx = -1
    for i, p in enumerate(paragraphs):
        if start_text in p.text:
            start_idx = i
        if end_text in p.text and start_idx != -1:
            end_idx = i
            break
    if start_idx != -1 and end_idx != -1:
        for i in range(end_idx, start_idx - 1, -1):
            p = paragraphs[i]._element
            p.getparent().remove(p)

# --- 儲蓄險特有邏輯 ---

def find_page_by_keyword(pdf_path, keyword):
    try:
        with pdfplumber.open(pdf_path) as pdf:
            for i, page in enumerate(pdf.pages):
                text = page.extract_text()
                if text and keyword in text:
                    return i + 1
    except Exception:
        pass
    return None

def get_value_by_text_search(pdf_path, page_num, keyword):
    try:
        with pdfplumber.open(pdf_path) as pdf:
            page = pdf.pages[page_num - 1]
            text = page.extract_text()
            if not text: return "N/A"
            lines = text.split('\n')
            for line in lines:
                if keyword in line:
                    matches = re.findall(r'[\d,]+', line)
                    nums = [m.replace(',', '').strip() for m in matches if m.replace(',', '').strip().isdigit()]
                    if nums: return nums[-1]
    except Exception:
        pass
    return "N/A"

def extract_values_from_filename_code1(filename):
    values = re.findall(r'\d+', filename)
    if len(values) >= 6:
        return values[:6]
    return None

def extract_nop_from_filename(filename):
    values = re.findall(r'\d+', filename)
    if len(values) >= 11:
        return values[5], values[7], values[10]
    return None, None, None

def extract_numeric_value_from_string(string):
    numbers = re.findall(r'\d+', string)
    return ''.join(numbers) if numbers else "N/A"

# --- Streamlit 界面 ---

st.set_page_config(page_title="PDF 計劃書自動化工具", layout="wide")
st.title("📄 PDF 計劃書自動化工具")

menu = ["儲蓄險", "儲蓄險添加", "一人重疾險", "二人重疾險", "三人重疾險", "四人重疾險"]
choice = st.sidebar.selectbox("選擇功能類型", menu)
template_file = st.sidebar.file_uploader("上傳 Word 模板 (.docx)", type=["docx"])

if choice in ["儲蓄險", "儲蓄險添加"]:
    pdf_file = st.file_uploader("選擇連續提取 PDF 文件", type=["pdf"])
    new_pdf_file = st.file_uploader("選擇分階段提取 PDF 文件 (可選)", type=["pdf"])
    
    if st.button("開始處理") and template_file and pdf_file:
        with st.spinner("正在處理中..."):
            with open("temp_pdf.pdf", "wb") as f:
                f.write(pdf_file.getbuffer())
            
            filename_values = extract_values_from_filename_code1(pdf_file.name)
            if not filename_values:
                st.error("PDF 文件名中未找到足夠的數值。")
            else:
                target_page = find_page_by_keyword("temp_pdf.pdf", "退保價值之説明摘要") or 6
                doc_fitz = fitz.open("temp_pdf.pdf")
                page_num_g_h = len(doc_fitz) - 6
                
                g = extract_table_value("temp_pdf.pdf", page_num_g_h, 11, 5)
                h = extract_table_value("temp_pdf.pdf", page_num_g_h, 12, 5)
                
                # s 的提取邏輯：與 g 同行 (11)，但取第一列 (0)
                s_raw = extract_table_value("temp_pdf.pdf", page_num_g_h, 11, 0)
                s = extract_numeric_value_from_string(s_raw)
                
                i = get_value_by_text_search("temp_pdf.pdf", target_page, "@ANB 56")
                j = get_value_by_text_search("temp_pdf.pdf", target_page, "@ANB 66")
                k = get_value_by_text_search("temp_pdf.pdf", target_page, "@ANB 76")
                l = get_value_by_text_search("temp_pdf.pdf", target_page, "@ANB 86")
                m = get_value_by_text_search("temp_pdf.pdf", target_page, "@ANB 96")
                
                st.write(f"### 提取數值驗證：")
                c1, c2, c3, c4, c5, c6 = st.columns(6)
                c1.metric("i (ANB 56)", i)
                c2.metric("j (ANB 66)", j)
                c3.metric("k (ANB 76)", k)
                c4.metric("l (ANB 86)", l)
                c5.metric("m (ANB 96)", m)
                c6.metric("s (年齡)", s)
                
                pdf_values = {"g": g, "h": h, "i": i, "j": j, "k": k, "l": l, "m": m, "s": s}
                values = dict(zip("abcdef", filename_values))
                values.update(pdf_values)
                
                remove_start, remove_end = None, None
                if new_pdf_file:
                    with open("temp_new_pdf.pdf", "wb") as f:
                        f.write(new_pdf_file.getbuffer())
                    n, o, p = extract_nop_from_filename(new_pdf_file.name)
                    new_doc_fitz = fitz.open("temp_new_pdf.pdf")
                    p_q_r = len(new_doc_fitz) - 6
                    q = extract_table_value("temp_new_pdf.pdf", p_q_r, 11, 5)
                    r = extract_table_value("temp_new_pdf.pdf", p_q_r, 12, 5)
                    s_new_raw = extract_table_value("temp_new_pdf.pdf", p_q_r, 11, 0)
                    s_new = extract_numeric_value_from_string(s_new_raw)
                    values.update({"n": n, "o": o, "p": p, "q": q, "r": r, "s": s_new})
                else:
                    remove_start = "在人生的重要阶段提取："
                    remove_end = "提取方式 3："
                
                output_bio = process_word_template(template_file, values, remove_start, remove_end)
                st.success("處理完成！")
                st.download_button("下載生成的 Word 文件", output_bio, file_name="output.docx")

elif choice in ["一人重疾險", "二人重疾險", "三人重疾險", "四人重疾險"]:
    num_files = {"一人重疾險": 1, "二人重疾險": 2, "三人重疾險": 3, "四人重疾險": 4}[choice]
    pdf_files = []
    for idx in range(num_files):
        pdf_files.append(st.file_uploader(f"選擇第 {idx+1} 個 PDF 文件", type=["pdf"], key=f"pdf_{idx}"))
    
    if st.button("開始處理") and template_file and all(pdf_files):
        with st.spinner("正在處理中..."):
            all_values = {}
            suffixes = ["", "1", "2", "3"]
            for idx, pdf in enumerate(pdf_files):
                suffix = suffixes[idx]
                temp_name = f"temp_pdf_{idx}.pdf"
                with open(temp_name, "wb") as f:
                    f.write(pdf.getbuffer())
                
                fn_vals = extract_values_from_filename(pdf.name)
                if fn_vals:
                    all_values.update(dict(zip([f"a{suffix}", f"b{suffix}", f"c{suffix}"], fn_vals)))
                
                # 嚴格對齊原始代碼邏輯
                d_vals = extract_row_values(temp_name, 3, "CIP2") or extract_row_values(temp_name, 3, "CIM3")
                d = d_vals[3] if len(d_vals) > 3 else "N/A"
                
                tables_p4 = camelot.read_pdf(temp_name, pages='4', flavor='stream')
                num_rows_p4 = tables_p4[0].df.shape[0] if tables_p4 else 0
                
                e = extract_table_value(temp_name, 4, num_rows_p4 - 8, 8)
                f = extract_table_value(temp_name, 4, num_rows_p4 - 6, 8)
                g = extract_table_value(temp_name, 4, num_rows_p4 - 4, 8)
                h = extract_table_value(temp_name, 4, num_rows_p4 - 2, 8)
                
                all_values.update({
                    f"d{suffix}": d, f"e{suffix}": e, f"f{suffix}": f, f"g{suffix}": g, f"h{suffix}": h
                })
            
            output_bio = process_word_template(template_file, all_values)
            st.success("處理完成！")
            st.download_button("下載生成的 Word 文件", output_bio, file_name="output.docx")

st.sidebar.markdown("---")
st.sidebar.info("請確保上傳的 PDF 格式與模板要求一致。")
