import streamlit as st
import os
import re
import camelot
import fitz  # PyMuPDF
import pdfplumber
from docx import Document
from docx.shared import Pt
from docx.oxml.ns import qn
from docx.enum.text import WD_LINE_SPACING
import pandas as pd
import io
import subprocess
import streamlit.components.v1 as components
import matplotlib.pyplot as plt
import base64
from jinja2 import Template
from html2image import Html2Image

# --- 移動端 App 化支持 (PWA) ---
pwa_html = """
<link rel="manifest" href="https://raw.githubusercontent.com/manus-agent/pwa-manifest/main/manifest.json">
<meta name="apple-mobile-web-app-capable" content="yes">
<meta name="apple-mobile-web-app-status-bar-style" content="black-translucent">
<meta name="apple-mobile-web-app-title" content="PDF工具">
<link rel="apple-touch-icon" href="https://cdn-icons-png.flaticon.com/512/4726/4726010.png">
<style>
    .stButton>button { width: 100%; border-radius: 10px; height: 3em; background-color: #007AFF; color: white; font-weight: bold; }
    .stMetric { background-color: #f0f2f6; padding: 10px; border-radius: 10px; margin-bottom: 10px; }
</style>
"""

# --- 公共函數部分 ---

def set_global_font_to_simsun(doc):
    """將文檔中所有文字設置為宋體，並取消對齊文檔網格以修復 PDF 行距問題"""
    for paragraph in doc.paragraphs:
        # 取消對齊文檔網格
        pPr = paragraph._element.get_or_add_pPr()
        snapToGrid = pPr.find(qn('w:snapToGrid'))
        if snapToGrid is not None:
            pPr.remove(snapToGrid)
        
        for run in paragraph.runs:
            run.font.name = 'SimSun'
            run._element.rPr.get_or_add_rFonts().set(qn('w:eastAsia'), 'SimSun')
            
    for table in doc.tables:
        for row in table.rows:
            for cell in row.cells:
                for paragraph in cell.paragraphs:
                    pPr = paragraph._element.get_or_add_pPr()
                    snapToGrid = pPr.find(qn('w:snapToGrid'))
                    if snapToGrid is not None:
                        pPr.remove(snapToGrid)
                    for run in paragraph.runs:
                        run.font.name = 'SimSun'
                        run._element.rPr.get_or_add_rFonts().set(qn('w:eastAsia'), 'SimSun')

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

def process_word_template(template_path, values, merge_start_text=None, merge_end_text=None, extra_removals=None):
    if not os.path.exists(template_path):
        raise FileNotFoundError(f"找不到模板文件: {template_path}")
    
    doc = Document(template_path)
    for paragraph in doc.paragraphs:
        replace_and_evaluate_in_paragraph(paragraph, values)
    for table in doc.tables:
        for row in table.rows:
            for cell in row.cells:
                for paragraph in cell.paragraphs:
                    replace_and_evaluate_in_paragraph(paragraph, values)
    
    if merge_start_text and merge_end_text:
        merge_paragraphs_and_delete_between_v2(doc, merge_start_text, merge_end_text)
    if extra_removals:
        for start, end in extra_removals:
            delete_specified_range(doc, start, end)
    
    set_global_font_to_simsun(doc)
            
    bio = io.BytesIO()
    doc.save(bio)
    bio.seek(0)
    return bio

def merge_paragraphs_and_delete_between_v2(doc, start_text, end_text):
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
        start_para = paragraphs[start_idx]
        end_para = paragraphs[end_idx]
        for run in start_para.runs:
            if start_text in run.text:
                run.text = run.text.replace(start_text, "")
        for run in end_para.runs:
            if end_text in run.text:
                run.text = run.text.replace(end_text, "")
        for run in end_para.runs:
            new_run = start_para.add_run(run.text)
            new_run.bold = run.bold
            new_run.italic = run.italic
            new_run.font.name = run.font.name
            new_run.font.size = run.font.size
            new_run.font.color.rgb = run.font.color.rgb
        for i in range(end_idx, start_idx, -1):
            p = paragraphs[i]._element
            p.getparent().remove(p)

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

def convert_docx_to_pdf(docx_bio):
    with open("temp_output.docx", "wb") as f:
        f.write(docx_bio.getbuffer())
    subprocess.run(["libreoffice", "--headless", "--convert-to", "pdf", "temp_output.docx"], check=True)
    with open("temp_output.pdf", "rb") as f:
        pdf_data = f.read()
    return pdf_data

# --- 營銷長圖生成邏輯 ---

def generate_marketing_image(values, html_template_path):
    # 1. 準備數據
    annual_premium = values.get('a', '10,000')
    pay_period = values.get('b', '5')
    total_invest = f"{float(annual_premium.replace(',','')) * int(pay_period):,.0f}"
    
    # 策略二動態數據
    strategy_b_items = []
    base_age = int(values.get('c', '29'))
    check_ages = [56, 66, 76, 86, 96]
    vals = [values.get('i'), values.get('j'), values.get('k'), values.get('l'), values.get('m')]
    
    max_val = 0
    for v in vals:
        if v and v != 'N/A':
            max_val = max(max_val, float(v.replace(',','')))

    for age, val in zip(check_ages, vals):
        if val and val != 'N/A':
            v_num = float(val.replace(',',''))
            ratio = v_num / (float(annual_premium.replace(',','')) * int(pay_period))
            width = (v_num / max_val * 100) if max_val > 0 else 0
            strategy_b_items.append({
                'age': f"{age}岁",
                'value': f"${add_thousand_separator(v_num)}",
                'ratio': f"{ratio:.1f}倍",
                'width': f"{width:.0f}%"
            })

    # 2. 生成圖表
    plt.figure(figsize=(6, 3), facecolor='white')
    x = [base_age + 5, 56, 66, 76, 86, 96]
    y = [0] + [float(v.replace(',','')) for v in vals if v != 'N/A']
    plt.fill_between(x, y, color='#fee2e2', alpha=0.5)
    plt.plot(x, y, color='#dc2626', linewidth=3)
    plt.axis('off')
    plt.tight_layout()
    
    img_bio = io.BytesIO()
    plt.savefig(img_bio, format='png', dpi=150, bbox_inches='tight')
    plt.close()
    chart_base64 = base64.b64encode(img_bio.getvalue()).decode()

    # 3. 渲染 HTML
    with open(html_template_path, 'r', encoding='utf-8') as f:
        html_content = f.read()
    
    # 簡單替換 (Jinja2 風格)
    template = Template(html_content)
    rendered_html = template.render(
        annual_premium=annual_premium,
        pay_period=pay_period,
        total_invest=total_invest,
        withdraw_period="39-90岁",
        withdraw_amount=values.get('h', '4,651'),
        legacy_value=values.get('m', '3,008,582'),
        strategy_b_items=strategy_b_items,
        chart_base64=chart_base64
    )
    
    # 4. 截圖
    hti = Html2Image(custom_flags=['--no-sandbox', '--disable-gpu', '--headless'])
    hti.screenshot(html_str=rendered_html, save_as='marketing.png', size=(420, 1200))
    
    with open('marketing.png', 'rb') as f:
        return f.read()

# --- 提取邏輯 ---

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

st.set_page_config(page_title="PDF 計劃書工具", layout="centered")
components.html(pwa_html, height=0)

st.title("📄 PDF 計劃書工具")

menu = ["儲蓄險", "儲蓄險添加", "一人重疾險", "二人重疾險", "三人重疾險", "四人重疾險"]
choice = st.selectbox("選擇功能類型", menu)

sub_choice = None
if "重疾險" in choice:
    sub_choice = st.radio("選擇產品類型", ["危疾單次保", "誠保一生"], horizontal=True)

export_options = ["Word (.docx)", "PDF (.pdf)"]
if choice in ["儲蓄險", "儲蓄險添加"]:
    export_options.append("營銷長圖 (.png)")
export_format = st.radio("選擇導出格式", export_options, horizontal=True)

with st.expander("📁 上傳 PDF 文件", expanded=True):
    if choice in ["儲蓄險", "儲蓄險添加"]:
        pdf_file = st.file_uploader("選擇連續提取 PDF", type=["pdf"])
        new_pdf_file = st.file_uploader("選擇分階段提取 PDF (可選)", type=["pdf"])
    else:
        num_files = {"一人重疾險": 1, "二人重疾險": 2, "三人重疾險": 3, "四人重疾險": 4}[choice]
        pdf_files = []
        for idx in range(num_files):
            pdf_files.append(st.file_uploader(f"選擇第 {idx+1} 個 PDF", type=["pdf"], key=f"pdf_{idx}"))

template_map = {
    "儲蓄險": "savings1.docx",
    "儲蓄險添加": "savings2.docx",
    "一人重疾險": {"危疾單次保": "one1.docx", "誠保一生": "one2.docx"},
    "二人重疾險": {"危疾單次保": "two1.docx", "誠保一生": "two2.docx"},
    "三人重疾險": {"危疾單次保": "three1.docx", "誠保一生": "three2.docx"},
    "四人重疾險": {"危疾單次保": "four1.docx", "誠保一生": "four2.docx"}
}

if st.button("🚀 開始處理"):
    with st.spinner("正在處理中..."):
        try:
            if "重疾險" in choice:
                template_path = template_map[choice][sub_choice]
            else:
                template_path = template_map[choice]
            
            if not os.path.exists(template_path) and "長圖" not in export_format:
                st.error(f"❌ 找不到模板文件: {template_path}。")
                st.stop()

            if choice in ["儲蓄險", "儲蓄險添加"]:
                if not pdf_file:
                    st.error("請上傳 PDF 文件！")
                else:
                    with open("temp_pdf.pdf", "wb") as f:
                        f.write(pdf_file.getbuffer())
                    filename_values = extract_values_from_filename_code1(pdf_file.name)
                    if not filename_values:
                        st.error("PDF 文件名格式不正確。")
                    else:
                        target_page = find_page_by_keyword("temp_pdf.pdf", "退保價值之説明摘要") or 6
                        doc_fitz = fitz.open("temp_pdf.pdf")
                        page_num_g_h = len(doc_fitz) - 6
                        g = extract_table_value("temp_pdf.pdf", page_num_g_h, 11, 5)
                        h = extract_table_value("temp_pdf.pdf", page_num_g_h, 12, 5)
                        s = extract_numeric_value_from_string(extract_table_value("temp_pdf.pdf", page_num_g_h, 11, 0))
                        i = get_value_by_text_search("temp_pdf.pdf", target_page, "@ANB 56")
                        j = get_value_by_text_search("temp_pdf.pdf", target_page, "@ANB 66")
                        k = get_value_by_text_search("temp_pdf.pdf", target_page, "@ANB 76")
                        l = get_value_by_text_search("temp_pdf.pdf", target_page, "@ANB 86")
                        m = get_value_by_text_search("temp_pdf.pdf", target_page, "@ANB 96")
                        pdf_values = {"g": g, "h": h, "i": i, "j": j, "k": k, "l": l, "m": m, "s": s}
                        values = dict(zip("abcdef", filename_values))
                        values.update(pdf_values)
                        
                        if "長圖" in export_format:
                            img_data = generate_marketing_image(values, "新建文本文档.html")
                            st.success("✅ 長圖生成完成！")
                            st.image(img_data)
                            st.download_button("📥 下載營銷長圖", img_data, file_name="概览.png", mime="image/png")
                        else:
                            merge_start, merge_end = None, None
                            extra_removals = []
                            if choice == "儲蓄險添加":
                                extra_removals.append(("信守明天多元货币储蓄计划概要：", "信守明天多元货币储蓄计划概要："))
                                extra_removals.append(("(保诚保险收益最高的储蓄产品，", "适合身体抱恙不能买寿险人士。"))
                            
                            if new_pdf_file:
                                with open("temp_new_pdf.pdf", "wb") as f:
                                    f.write(new_pdf_file.getbuffer())
                                n, o, p = extract_nop_from_filename(new_pdf_file.name)
                                new_doc_fitz = fitz.open("temp_new_pdf.pdf")
                                p_q_r = len(new_doc_fitz) - 6
                                q = extract_table_value("temp_new_pdf.pdf", p_q_r, 11, 5)
                                r = extract_table_value("temp_new_pdf.pdf", p_q_r, 12, 5)
                                s_new = extract_numeric_value_from_string(extract_table_value("temp_new_pdf.pdf", p_q_r, 11, 0))
                                values.update({"n": n, "o": o, "p": p, "q": q, "r": r, "s": s_new})
                            else:
                                merge_start = "在人生的重要阶段提取："
                                merge_end = "提取方式 3："
                            
                            output_docx = process_word_template(template_path, values, merge_start, merge_end, extra_removals)
                            if "PDF" in export_format:
                                pdf_data = convert_docx_to_pdf(output_docx)
                                st.success("✅ PDF 生成完成！")
                                st.download_button("📥 下載 PDF 文件", pdf_data, file_name="概览.pdf", mime="application/pdf")
                            else:
                                st.success("✅ Word 生成完成！")
                                st.download_button("📥 下載 Word 文件", output_docx, file_name="概览.docx")
                        
            elif "重疾險" in choice:
                if not all(pdf_files):
                    st.error("請上傳所有 PDF 文件！")
                else:
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
                        target_page_summary = find_page_by_keyword(temp_name, "説明摘要") or 6
                        e = get_value_by_text_search(temp_name, target_page_summary, "@ANB 56")
                        f = get_value_by_text_search(temp_name, target_page_summary, "@ANB 66")
                        g = get_value_by_text_search(temp_name, target_page_summary, "@ANB 76")
                        h = get_value_by_text_search(temp_name, target_page_summary, "@ANB 86")
                        target_page_d = find_page_by_keyword(temp_name, "建議書摘要") or 5
                        d_vals = extract_row_values(temp_name, target_page_d, "CIP2") or extract_row_values(temp_name, target_page_d, "CIM3")
                        d = d_vals[3] if len(d_vals) > 3 else "N/A"
                        all_values.update({f"d{suffix}": d, f"e{suffix}": e, f"f{suffix}": f, f"g{suffix}": g, f"h{suffix}": h})
                    
                    output_docx = process_word_template(template_path, all_values)
                    if "PDF" in export_format:
                        pdf_data = convert_docx_to_pdf(output_docx)
                        st.success("✅ PDF 生成完成！")
                        st.download_button("📥 下載 PDF 文件", pdf_data, file_name="概览.pdf", mime="application/pdf")
                    else:
                        st.success("✅ Word 生成完成！")
                        st.download_button("📥 下載 Word 文件", output_docx, file_name="概览.docx")

        except Exception as e:
            st.error(f"❌ 發生錯誤: {str(e)}")

st.markdown("---")
st.caption("💡 提示：請確保所有 Word 模板和 HTML 模板已上傳至 GitHub 倉庫根目錄。")
