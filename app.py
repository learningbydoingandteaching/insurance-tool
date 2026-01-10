import streamlit as st
import os
import re
import camelot
import fitz  # PyMuPDF
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
    except Exception as e:
        st.error(f"提取表格數值出錯: {e}")
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
    except Exception as e:
        st.error(f"提取行數值出錯: {e}")
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
        # 安全評估簡單數學表達式
        result = eval(expression, {"__builtins__": None}, {})
        return add_thousand_separator(result)
    except Exception as e:
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
    
    # 處理段落
    for paragraph in doc.paragraphs:
        replace_and_evaluate_in_paragraph(paragraph, values)
    
    # 處理表格
    for table in doc.tables:
        for row in table.rows:
            for cell in row.cells:
                for paragraph in cell.paragraphs:
                    replace_and_evaluate_in_paragraph(paragraph, values)
    
    # 處理刪除邏輯
    if remove_text_start and remove_text_end:
        delete_specified_runs(doc, remove_text_start, remove_text_end)
        
    # 保存到內存
    bio = io.BytesIO()
    doc.save(bio)
    bio.seek(0)
    return bio

def delete_specified_runs(doc, start_text, end_text):
    inside_delete_range = False
    runs_to_delete = []
    paragraphs_to_check = set()

    for paragraph in doc.paragraphs:
        for run in paragraph.runs:
            if start_text in run.text:
                inside_delete_range = True
            if inside_delete_range:
                runs_to_delete.append(run)
                paragraphs_to_check.add(paragraph)
            if end_text in run.text:
                inside_delete_range = False
                for run_to_delete in runs_to_delete[:-1]:
                    run_to_delete.clear()
                runs_to_delete = []
                paragraphs_to_check.add(paragraph)
                break

    for paragraph in list(paragraphs_to_check):
        if not paragraph.text.strip():
            p = paragraph._element
            p.getparent().remove(p)

# --- 儲蓄險特有邏輯 ---

def extract_values_from_filename_code1(filename):
    values = re.findall(r'\d+', filename)
    if len(values) >= 6:
        return values[:6]
    return None

def extract_nop_from_filename(filename):
    values = re.findall(r'\d+', filename)
    if len(values) >= 11:
        n = values[5]
        o = values[7]
        p = values[10]
        return n, o, p
    return None, None, None

def extract_numeric_value_from_string(string):
    numbers = re.findall(r'\d+', string)
    return ''.join(numbers) if numbers else "N/A"

def get_i_j_k_l_m_from_page6(pdf_path):
    try:
        tables = camelot.read_pdf(pdf_path, pages='6', flavor='stream')
        if len(tables) > 0:
            df = tables[0].df
            num_rows = df.shape[0]
            
            def get_val(row_from_bottom):
                try:
                    idx = num_rows - row_from_bottom
                    val = df.iat[idx, -1]
                    return val.replace(',', '').replace(' ', '')
                except:
                    return "N/A"
            
            i = get_val(10) # ANB 56
            j = get_val(8)  # ANB 66
            k = get_val(6)  # ANB 76
            l = get_val(4)  # ANB 86
            m = get_val(2)  # ANB 96
            return i, j, k, l, m
    except Exception as e:
        st.error(f"提取 i,j,k,l,m 出錯: {e}")
    return "N/A", "N/A", "N/A", "N/A", "N/A"

# --- Streamlit 界面 ---

st.set_page_config(page_title="PDF 計劃書自動化工具", layout="wide")
st.title("📄 PDF 計劃書自動化工具")

menu = ["儲蓄險", "儲蓄險添加", "二人重疾險", "三人重疾險", "四人重疾險"]
choice = st.sidebar.selectbox("選擇功能類型", menu)

template_file = st.sidebar.file_uploader("上傳 Word 模板 (.docx)", type=["docx"])

if choice in ["儲蓄險", "儲蓄险添加"]:
    pdf_file = st.file_uploader("選擇連續提取 PDF 文件", type=["pdf"])
    new_pdf_file = st.file_uploader("選擇分階段提取 PDF 文件 (可選)", type=["pdf"])
    
    if st.button("開始處理") and template_file and pdf_file:
        with st.spinner("正在處理中..."):
            # 保存臨時文件
            with open("temp_pdf.pdf", "wb") as f:
                f.write(pdf_file.getbuffer())
            
            filename_values = extract_values_from_filename_code1(pdf_file.name)
            if not filename_values:
                st.error("PDF 文件名中未找到足夠的數值。")
            else:
                # 提取 g, h
                doc_fitz = fitz.open("temp_pdf.pdf")
                total_pages = len(doc_fitz)
                page_num_g_h = total_pages - 6
                g = extract_table_value("temp_pdf.pdf", page_num_g_h, 11, 5)
                h = extract_table_value("temp_pdf.pdf", page_num_g_h, 12, 5)
                
                # 提取 i, j, k, l, m
                i, j, k, l, m = get_i_j_k_l_m_from_page6("temp_pdf.pdf")
                
                pdf_values = {"g": g, "h": h, "i": i, "j": j, "k": k, "l": l, "m": m}
                
                # 提取 s (如果是 code4)
                if choice == "儲蓄險添加":
                    s_string = extract_table_value("temp_pdf.pdf", page_num_g_h, 11, 0)
                    pdf_values["s"] = extract_numeric_value_from_string(s_string)
                
                values = dict(zip("abcdef", filename_values))
                values.update(pdf_values)
                
                remove_start = None
                remove_end = None
                
                if new_pdf_file:
                    with open("temp_new_pdf.pdf", "wb") as f:
                        f.write(new_pdf_file.getbuffer())
                    n, o, p = extract_nop_from_filename(new_pdf_file.name)
                    new_doc_fitz = fitz.open("temp_new_pdf.pdf")
                    page_num_q_r = len(new_doc_fitz) - 6
                    q = extract_table_value("temp_new_pdf.pdf", page_num_q_r, 11, 5)
                    r = extract_table_value("temp_new_pdf.pdf", page_num_q_r, 12, 5)
                    s_new = extract_numeric_value_from_string(extract_table_value("temp_new_pdf.pdf", page_num_q_r, 11, 0))
                    values.update({"n": n, "o": o, "p": p, "q": q, "r": r, "s": s_new})
                else:
                    remove_start = "在人生的重要階段提取："
                    remove_end = "不提取分紅，在某年，把累積的本金"
                
                # 處理 Word
                output_bio = process_word_template(template_file, values, remove_start, remove_end)
                st.success("處理完成！")
                st.download_button("下載生成的 Word 文件", output_bio, file_name="output.docx")

elif choice in ["二人重疾險", "三人重疾險", "四人重疾險"]:
    num_files = {"二人重疾險": 2, "三人重疾險": 3, "四人重疾險": 4}[choice]
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
