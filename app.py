import streamlit as st
import os
import re
import camelot
import fitz  # PyMuPDF
from docx import Document
import copy
import tempfile
import shutil

# ==========================================
# 核心逻辑函数 (源自您的原始代码，去除了Tkinter)
# ==========================================

def extract_values_from_filename(filename):
    values = re.findall(r'\d+', filename)
    if len(values) >= 3:
        return values[:3]
    return None

def extract_table_value(pdf_path, page_num, row_num, col_num):
    # Camelot 需要物理路径
    tables = camelot.read_pdf(pdf_path, pages=str(page_num), flavor='stream')
    for table in tables:
        df = table.df
        try:
            value = df.iat[int(row_num), int(col_num)].replace(',', '')
            return value
        except IndexError:
            continue
    return "N/A"

def extract_row_values(pdf_path, page_num, keyword):
    tables = camelot.read_pdf(pdf_path, pages=str(page_num), flavor='stream')
    for table in tables:
        df = table.df
        for i, row in df.iterrows():
            if keyword in row.to_string():
                values = [val.replace(',', '') for val in re.findall(r"[\d,.]+", row.to_string())]
                return values
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
        result = eval(expression)
        return add_thousand_separator(result)
    except Exception as e:
        print(f"计算表达式时出错: {expression}. 错误信息: {e}")
        return "N/A"

def replace_and_evaluate_in_run(run, values):
    full_text = run.text
    for key, value in values.items():
        placeholder = f"{{{key}}}"
        full_text = full_text.replace(placeholder, value if value is not None else "N/A")

    expressions = re.findall(r'\{\{[^\}]+\}\}', full_text)
    for expr in expressions:
        expr_clean = expr.strip("{}")
        result = evaluate_expression(expr_clean, values)
        full_text = full_text.replace(expr, result)

    run.text = full_text

def replace_and_evaluate_in_paragraph(paragraph, values):
    for run in paragraph.runs:
        replace_and_evaluate_in_run(run, values)

def replace_values_in_word_template(template_path, output_path, values):
    doc = Document(template_path)
    for paragraph in doc.paragraphs:
        replace_and_evaluate_in_paragraph(paragraph, values)
    for table in doc.tables:
        for row in table.rows:
            for cell in row.cells:
                for paragraph in cell.paragraphs:
                    replace_and_evaluate_in_paragraph(paragraph, values)
    doc.save(output_path)

# --- 储蓄险专用函数 ---

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

    for paragraph in paragraphs_to_check:
        if not paragraph.text.strip():
            p = paragraph._element
            p.getparent().remove(p)
            p._element = None

def extract_numeric_value_from_string(string):
    numbers = re.findall(r'\d+', string)
    return ''.join(numbers) if numbers else "N/A"

def replace_values_in_word_template_with_delete(template_path, output_path, values, remove_text_start, remove_text_end):
    doc = Document(template_path)
    for paragraph in doc.paragraphs:
        replace_and_evaluate_in_paragraph(paragraph, values)
    for table in doc.tables:
        for row in table.rows:
            for cell in row.cells:
                for paragraph in cell.paragraphs:
                    replace_and_evaluate_in_paragraph(paragraph, values)
    if remove_text_start and remove_text_end:
        delete_specified_runs(doc, remove_text_start, remove_text_end)
    doc.save(output_path)

def replace_values_in_word_template_append(template_path, output_path, values, remove_text_start=None, remove_text_end=None):
    # 注意：在Code4中，如果output_path存在，则读取它；但在Web版中，output_path是新生成的
    # 这里的逻辑稍微调整：Web版每次都是生成新文件，所以我们假设 template_path 就是基础文件
    
    # 为了兼容原逻辑，我们直接操作 template_path 对应的文档对象
    template_doc = Document(template_path)

    for paragraph in template_doc.paragraphs:
        replace_and_evaluate_in_paragraph(paragraph, values)
    for table in template_doc.tables:
        for row in table.rows:
            for cell in row.cells:
                for paragraph in cell.paragraphs:
                    replace_and_evaluate_in_paragraph(paragraph, values)
    if remove_text_start and remove_text_end:
        delete_specified_runs(template_doc, remove_text_start, remove_text_end)

    # 保存
    template_doc.save(output_path)


# ==========================================
# 业务逻辑处理函数 (修改为抛出异常而非弹窗)
# ==========================================

def process_code1(pdf_file, new_pdf_file, template_path, output_path):
    pdf_filename = os.path.basename(pdf_file)
    filename_values = extract_values_from_filename_code1(pdf_filename)
    if not filename_values:
        raise Exception("PDF 文件名中未找到足够的数值 (需要至少6个数字)。")

    doc = fitz.open(pdf_file)
    total_pages = len(doc)
    page_num_g_h = total_pages - 6

    g = extract_table_value(pdf_file, page_num_g_h, 11, 5)
    h = extract_table_value(pdf_file, page_num_g_h, 12, 5)

    # 提取第6页
    tables_page_6 = camelot.read_pdf(pdf_file, pages='6', flavor='stream')
    if len(tables_page_6) > 0:
        df_page_6 = tables_page_6[0].df
        num_rows_page_6 = df_page_6.shape[0]

        def get_val_from_last_col(row_from_bottom):
            try:
                target_row_idx = num_rows_page_6 - row_from_bottom
                val = df_page_6.iat[target_row_idx, -2]
                return val.replace(',', '').replace(' ', '')
            except Exception as e:
                return "N/A"

        i = get_val_from_last_col(10)
        j = get_val_from_last_col(8)
        k = get_val_from_last_col(6)
        l = get_val_from_last_col(4)
        m = get_val_from_last_col(2)
    else:
        i = j = k = l = m = "N/A"

    pdf_values = {"g": g, "h": h, "i": i, "j": j, "k": k, "l": l, "m": m}
    values = dict(zip("abcdef", filename_values))
    values.update(pdf_values)

    if not new_pdf_file:
        remove_text_start = "在人生的重要阶段提取："
        remove_text_end = "不提取分红，在某年，把累积的本金"
        replace_values_in_word_template_with_delete(template_path, output_path, values, remove_text_start, remove_text_end)
        return "处理完成 (单PDF模式)"

    # 处理第二个PDF
    new_pdf_filename = os.path.basename(new_pdf_file)
    n, o, p = extract_nop_from_filename(new_pdf_filename)
    if not n or not o or not p:
        raise Exception("第二个 PDF 文件名中未找到足够的数值用于 n, o, p。")

    new_doc = fitz.open(new_pdf_file)
    total_new_pages = len(new_doc)
    page_num_q_r = total_new_pages - 6

    q = extract_table_value(new_pdf_file, page_num_q_r, 11, 5)
    r = extract_table_value(new_pdf_file, page_num_q_r, 12, 5)
    s_string = extract_table_value(new_pdf_file, page_num_q_r, 11, 0)
    s = extract_numeric_value_from_string(s_string)

    new_pdf_values = {"n": n, "o": o, "p": p, "q": q, "r": r, "s": s}
    values.update(new_pdf_values)

    replace_values_in_word_template(template_path, output_path, values)
    return "处理完成 (双PDF模式)"


def process_code4(pdf_file, new_pdf_file, template_path, output_path):
    # 逻辑与Code1类似，但使用 append 模式
    pdf_filename = os.path.basename(pdf_file)
    filename_values = extract_values_from_filename_code1(pdf_filename)
    if not filename_values:
        raise Exception("PDF 文件名中未找到足够的数值。")

    doc = fitz.open(pdf_file)
    total_pages = len(doc)
    page_num_g_h = total_pages - 6

    g = extract_table_value(pdf_file, page_num_g_h, 11, 5)
    h = extract_table_value(pdf_file, page_num_g_h, 12, 5)
    
    page_num_s = total_pages - 6
    s_string = extract_table_value(pdf_file, page_num_s, 11, 0)
    s = extract_numeric_value_from_string(s_string)

    tables_page_6 = camelot.read_pdf(pdf_file, pages='6', flavor='stream')
    i = j = k = l = m = "N/A"
    if len(tables_page_6) > 0:
        df_page_6 = tables_page_6[0].df
        num_rows_page_6 = df_page_6.shape[0]
        def get_val_from_last_col(row_from_bottom):
            try:
                target_row_idx = num_rows_page_6 - row_from_bottom
                val = df_page_6.iat[target_row_idx, -2]
                return val.replace(',', '').replace(' ', '')
            except Exception: return "N/A"
        i = get_val_from_last_col(10)
        j = get_val_from_last_col(8)
        k = get_val_from_last_col(6)
        l = get_val_from_last_col(4)
        m = get_val_from_last_col(2)

    pdf_values = {"g": g, "h": h, "i": i, "j": j, "k": k, "l": l, "m": m, "s": s}
    values = dict(zip("abcdef", filename_values))
    values.update(pdf_values)

    if not new_pdf_file:
        remove_text_start = "在人生的重要阶段提取："
        remove_text_end = "不提取分红，在某年，把累积的本金"
        replace_values_in_word_template_append(template_path, output_path, values, remove_text_start, remove_text_end)
        return "储蓄险添加处理完成 (单PDF)"

    new_pdf_filename = os.path.basename(new_pdf_file)
    n, o, p = extract_nop_from_filename(new_pdf_filename)
    if not n or not o or not p:
        raise Exception("第二个 PDF 文件名中未找到足够的数值用于 n, o, p。")

    new_doc = fitz.open(new_pdf_file)
    total_new_pages = len(new_doc)
    page_num_q_r = total_new_pages - 6
    q = extract_table_value(new_pdf_file, page_num_q_r, 11, 5)
    r = extract_table_value(new_pdf_file, page_num_q_r, 12, 5)

    new_pdf_values = {"n": n, "o": o, "p": p, "q": q, "r": r}
    values.update(new_pdf_values)

    replace_values_in_word_template_append(template_path, output_path, values)
    return "储蓄险添加处理完成 (双PDF)"


def process_ci_common(pdf_files, template_path, output_path):
    # 通用的重疾险处理逻辑 (1-4人)
    # pdf_files 是一个列表
    
    all_values = {}
    
    for idx, pdf_file in enumerate(pdf_files):
        suffix = "" if idx == 0 else str(idx) # 第一个人无后缀，第二个是1，第三个是2...
        if idx == 0: suffix_keys = ["a", "b", "c"]
        else: suffix_keys = [f"a{idx}", f"b{idx}", f"c{idx}"]
        
        pdf_filename = os.path.basename(pdf_file)
        filename_values = extract_values_from_filename(pdf_filename)
        if not filename_values:
            raise Exception(f"第 {idx+1} 个 PDF 文件名中未找到足够的数值。")
            
        # 提取数据
        d_values = extract_row_values(pdf_file, 3, "CIP2") or extract_row_values(pdf_file, 3, "CIM3")
        d = d_values[3] if len(d_values) > 3 else "N/A"

        num_rows_page_4 = 0
        tables_page_4 = camelot.read_pdf(pdf_file, pages='4', flavor='stream')
        for table in tables_page_4:
            df_page_4 = table.df
            num_rows_page_4 = df_page_4.shape[0]

        e = extract_table_value(pdf_file, 4, num_rows_page_4 - 8, 8)
        f = extract_table_value(pdf_file, 4, num_rows_page_4 - 6, 8)
        g = extract_table_value(pdf_file, 4, num_rows_page_4 - 4, 8)
        h = extract_table_value(pdf_file, 4, num_rows_page_4 - 2, 8)

        key_d = "d" + ("" if idx == 0 else str(idx))
        key_e = "e" + ("" if idx == 0 else str(idx))
        key_f = "f" + ("" if idx == 0 else str(idx))
        key_g = "g" + ("" if idx == 0 else str(idx))
        key_h = "h" + ("" if idx == 0 else str(idx))

        pdf_values = {
            key_d: d, key_e: e, key_f: f, key_g: g, key_h: h
        }
        
        all_values.update(dict(zip(suffix_keys, filename_values)))
        all_values.update(pdf_values)

    replace_values_in_word_template(template_path, output_path, all_values)
    return f"重疾险 ({len(pdf_files)}人) 处理完成"


# ==========================================
# Streamlit 界面部分
# ==========================================

st.set_page_config(page_title="保险计划书生成器", layout="wide")

st.title("📋 保险计划书自动生成器")
st.markdown("---")

# 侧边栏选择模式
mode = st.sidebar.radio(
    "请选择功能模式",
    [
        "储蓄险 (Code1)",
        "储蓄险-添加模式 (Code4)",
        "一人重疾险 (Code2)",
        "二人重疾险 (Code5)",
        "三人重疾险 (Code6)",
        "四人重疾险 (Code7)"
    ]
)

st.header(f"当前模式: {mode}")

# 文件上传区
uploaded_pdfs = []
uploaded_template = st.file_uploader("上传 Word 模板 (.docx)", type=["docx"])

# 根据模式显示不同的 PDF 上传框
if "储蓄险" in mode:
    pdf1 = st.file_uploader("上传主 PDF 文件", type=["pdf"], key="s1")
    pdf2 = st.file_uploader("上传第二个 PDF 文件 (可选)", type=["pdf"], key="s2")
    if pdf1: uploaded_pdfs.append(pdf1)
    if pdf2: uploaded_pdfs.append(pdf2)
else:
    # 重疾险
    count = 1
    if "二人" in mode: count = 2
    if "三人" in mode: count = 3
    if "四人" in mode: count = 4
    
    for i in range(count):
        pdf = st.file_uploader(f"上传第 {i+1} 个人的 PDF", type=["pdf"], key=f"ci_{i}")
        if pdf: uploaded_pdfs.append(pdf)

# 开始生成按钮
if st.button("🚀 开始生成", type="primary"):
    if not uploaded_template:
        st.error("请上传 Word 模板文件！")
    elif len(uploaded_pdfs) == 0:
        st.error("请至少上传一个 PDF 文件！")
    else:
        # 创建临时目录
        with tempfile.TemporaryDirectory() as temp_dir:
            try:
                # 1. 保存 Word 模板
                temp_tpl_path = os.path.join(temp_dir, uploaded_template.name)
                with open(temp_tpl_path, "wb") as f:
                    f.write(uploaded_template.getvalue())
                
                # 2. 保存 PDF 文件 (保持原始文件名，这对您的正则逻辑至关重要)
                pdf_paths = []
                for up_pdf in uploaded_pdfs:
                    p_path = os.path.join(temp_dir, up_pdf.name)
                    with open(p_path, "wb") as f:
                        f.write(up_pdf.getvalue())
                    pdf_paths.append(p_path)

                output_path = os.path.join(temp_dir, "generated_plan.docx")
                result_msg = ""

                # 3. 调用逻辑
                with st.spinner("正在分析数据并生成文档..."):
                    if "Code1" in mode:
                        p2 = pdf_paths[1] if len(pdf_paths) > 1 else None
                        result_msg = process_code1(pdf_paths[0], p2, temp_tpl_path, output_path)
                    
                    elif "Code4" in mode:
                        p2 = pdf_paths[1] if len(pdf_paths) > 1 else None
                        result_msg = process_code4(pdf_paths[0], p2, temp_tpl_path, output_path)
                    
                    else:
                        # 重疾险系列 (Code2, 5, 6, 7)
                        # 检查文件数量是否匹配
                        expected_count = 1
                        if "二人" in mode: expected_count = 2
                        if "三人" in mode: expected_count = 3
                        if "四人" in mode: expected_count = 4
                        
                        if len(pdf_paths) != expected_count:
                            raise Exception(f"当前模式需要 {expected_count} 个PDF文件，但您上传了 {len(pdf_paths)} 个。")
                        
                        result_msg = process_ci_common(pdf_paths, temp_tpl_path, output_path)

                # 4. 成功后显示下载按钮
                st.success(f"✅ {result_msg}")
                
                with open(output_path, "rb") as f:
                    st.download_button(
                        label="📥 下载生成的计划书",
                        data=f,
                        file_name="保险计划书_生成版.docx",
                        mime="application/vnd.openxmlformats-officedocument.wordprocessingml.document"
                    )

            except Exception as e:
                st.error(f"❌ 发生错误: {str(e)}")
                st.info("提示: 请确保 PDF 文件名包含所需的数字编号，且格式正确。")
