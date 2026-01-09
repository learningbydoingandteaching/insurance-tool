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
# 核心逻辑函数
# ==========================================

def extract_values_from_filename(filename):
    values = re.findall(r'\d+', filename)
    if len(values) >= 3:
        return values[:3]
    return None

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

def extract_numeric_value_from_string(string):
    numbers = re.findall(r'\d+', string)
    return ''.join(numbers) if numbers else "N/A"

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

# ==========================================
# Streamlit 界面逻辑
# ==========================================

st.set_page_config(page_title="保险计划书生成工具", layout="wide")

st.title("📋 保险计划书自动生成工具")
st.markdown("---")

# 侧边栏：选择功能
option = st.sidebar.selectbox(
    "请选择功能模式",
    ("储蓄险 (Code 1)", "一人重疾险 (Code 2)", "二人重疾险 (Code 5)")
)

# 临时文件夹管理
if 'temp_dir' not in st.session_state:
    st.session_state.temp_dir = tempfile.mkdtemp()

def save_uploaded_file(uploaded_file):
    if uploaded_file is not None:
        file_path = os.path.join(st.session_state.temp_dir, uploaded_file.name)
        with open(file_path, "wb") as f:
            f.write(uploaded_file.getbuffer())
        return file_path
    return None

# -------------------------------------------------------
# 模式 1: 储蓄险 (Code 1)
# -------------------------------------------------------
if option == "储蓄险 (Code 1)":
    st.header("储蓄险计划书生成")
    
    col1, col2 = st.columns(2)
    with col1:
        pdf_file = st.file_uploader("上传主 PDF 文件 (包含 a,b,c,d,e,f)", type=["pdf"], key="s_pdf")
        new_pdf_file = st.file_uploader("上传分阶段提取 PDF (可选, 包含 n,o,p)", type=["pdf"], key="s_new_pdf")
    with col2:
        template_file = st.file_uploader("上传 Word 模板 (.docx)", type=["docx"], key="s_word")

    if st.button("开始生成", key="btn_s"):
        if pdf_file and template_file:
            try:
                with st.spinner("正在处理..."):
                    # 保存文件到本地临时目录
                    pdf_path = save_uploaded_file(pdf_file)
                    template_path = save_uploaded_file(template_file)
                    new_pdf_path = save_uploaded_file(new_pdf_file) if new_pdf_file else None
                    
                    output_filename = f"Generated_{os.path.splitext(template_file.name)[0]}.docx"
                    output_path = os.path.join(st.session_state.temp_dir, output_filename)

                    # --- 提取逻辑开始 ---
                    pdf_filename = os.path.basename(pdf_path)
                    filename_values = extract_values_from_filename_code1(pdf_filename)
                    
                    if not filename_values:
                        st.error("主 PDF 文件名格式错误，未找到足够的数值 (a-f)。")
                    else:
                        # 提取 g, h
                        doc = fitz.open(pdf_path)
                        total_pages = len(doc)
                        page_num_g_h = total_pages - 6
                        g = extract_table_value(pdf_path, page_num_g_h, 11, 5)
                        h = extract_table_value(pdf_path, page_num_g_h, 12, 5)

                        # --- 【关键修改点】提取 i, j, k, l, m ---
                        # 读取PDF第6页
                        tables_page_6 = camelot.read_pdf(pdf_path, pages='6', flavor='stream')
                        
                        if len(tables_page_6) > 0:
                            df_page_6 = tables_page_6[0].df
                            num_rows_page_6 = df_page_6.shape[0]

                            # 定义内部函数
                            def get_val_from_target_col(row_from_bottom):
                                try:
                                    target_row_idx = num_rows_page_6 - row_from_bottom
                                    # 【此处已修改为 -2】
                                    val = df_page_6.iat[target_row_idx, -2] 
                                    return val.replace(',', '').replace(' ', '')
                                except Exception as e:
                                    return "N/A"

                            i = get_val_from_target_col(10) # ANB 56
                            j = get_val_from_target_col(8)  # ANB 66
                            k = get_val_from_target_col(6)  # ANB 76
                            l = get_val_from_target_col(4)  # ANB 86
                            m = get_val_from_target_col(2)  # ANB 96
                        else:
                            i = j = k = l = m = "N/A"
                        # -------------------------------------

                        pdf_values = {
                            "g": g, "h": h, "i": i, "j": j, "k": k, "l": l, "m": m
                        }
                        values = dict(zip("abcdef", filename_values))
                        values.update(pdf_values)

                        # 分支：是否有第二个PDF
                        if not new_pdf_path:
                            remove_text_start = "在人生的重要阶段提取："
                            remove_text_end = "不提取分红，在某年，把累积的本金"
                            replace_values_in_word_template_with_delete(template_path, output_path, values, remove_text_start, remove_text_end)
                        else:
                            new_pdf_filename = os.path.basename(new_pdf_path)
                            n, o, p = extract_nop_from_filename(new_pdf_filename)
                            if not n or not o or not p:
                                st.warning("第二个 PDF 文件名未找到 n, o, p，将跳过相关替换。")
                            else:
                                new_doc = fitz.open(new_pdf_path)
                                total_new_pages = len(new_doc)
                                page_num_q_r = total_new_pages - 6
                                q = extract_table_value(new_pdf_path, page_num_q_r, 11, 5)
                                r = extract_table_value(new_pdf_path, page_num_q_r, 12, 5)
                                s_string = extract_table_value(new_pdf_path, page_num_q_r, 11, 0)
                                s = extract_numeric_value_from_string(s_string)

                                new_pdf_values = {"n": n, "o": o, "p": p, "q": q, "r": r, "s": s}
                                values.update(new_pdf_values)
                                replace_values_in_word_template(template_path, output_path, values)

                        st.success("✅ 生成成功！")
                        with open(output_path, "rb") as file:
                            st.download_button(
                                label="下载生成的 Word 文档",
                                data=file,
                                file_name=output_filename,
                                mime="application/vnd.openxmlformats-officedocument.wordprocessingml.document"
                            )

            except Exception as e:
                st.error(f"发生错误: {e}")
        else:
            st.warning("请上传必要的文件 (主 PDF 和 Word 模板)。")

# -------------------------------------------------------
# 模式 2: 一人重疾险 (Code 2)
# -------------------------------------------------------
elif option == "一人重疾险 (Code 2)":
    st.header("一人重疾险计划书生成")
    
    col1, col2 = st.columns(2)
    with col1:
        pdf_file = st.file_uploader("上传 PDF 文件", type=["pdf"], key="c2_pdf")
    with col2:
        template_file = st.file_uploader("上传 Word 模板 (.docx)", type=["docx"], key="c2_word")

    if st.button("开始生成", key="btn_c2"):
        if pdf_file and template_file:
            try:
                with st.spinner("正在处理..."):
                    pdf_path = save_uploaded_file(pdf_file)
                    template_path = save_uploaded_file(template_file)
                    
                    output_filename = f"Generated_{os.path.splitext(template_file.name)[0]}.docx"
                    output_path = os.path.join(st.session_state.temp_dir, output_filename)

                    # 提取逻辑
                    pdf_filename = os.path.basename(pdf_path)
                    filename_values = extract_values_from_filename(pdf_filename)
                    
                    if not filename_values:
                        st.error("PDF 文件名中未找到足够的数值 (a, b, c)。")
                    else:
                        # 提取 d (CIP2 或 CIM3)
                        def extract_row_values_local(pdf_path, page_num, keyword):
                            tables = camelot.read_pdf(pdf_path, pages=str(page_num), flavor='stream')
                            for table in tables:
                                df = table.df
                                for i, row in df.iterrows():
                                    if keyword in row.to_string():
                                        values = [val.replace(',', '') for val in re.findall(r"[\d,.]+", row.to_string())]
                                        return values
                            return []

                        d_values = extract_row_values_local(pdf_path, 3, "CIP2") or extract_row_values_local(pdf_path, 3, "CIM3")
                        d = d_values[3] if len(d_values) > 3 else "N/A"

                        # 提取 e, f, g, h (Page 4)
                        num_rows_page_4 = 0
                        tables_page_4 = camelot.read_pdf(pdf_path, pages='4', flavor='stream')
                        for table in tables_page_4:
                            df_page_4 = table.df
                            num_rows_page_4 = df_page_4.shape[0]

                        e = extract_table_value(pdf_path, 4, num_rows_page_4 - 8, 8)
                        f = extract_table_value(pdf_path, 4, num_rows_page_4 - 6, 8)
                        g = extract_table_value(pdf_path, 4, num_rows_page_4 - 4, 8)
                        h = extract_table_value(pdf_path, 4, num_rows_page_4 - 2, 8)

                        pdf_values = {"d": d, "e": e, "f": f, "g": g, "h": h}
                        values = dict(zip("abc", filename_values))
                        values.update(pdf_values)

                        replace_values_in_word_template(template_path, output_path, values)

                        st.success("✅ 生成成功！")
                        with open(output_path, "rb") as file:
                            st.download_button(
                                label="下载生成的 Word 文档",
                                data=file,
                                file_name=output_filename,
                                mime="application/vnd.openxmlformats-officedocument.wordprocessingml.document"
                            )
            except Exception as e:
                st.error(f"发生错误: {e}")
        else:
            st.warning("请上传必要的文件。")

# -------------------------------------------------------
# 模式 3: 二人重疾险 (Code 5)
# -------------------------------------------------------
elif option == "二人重疾险 (Code 5)":
    st.header("二人重疾险计划书生成")
    
    col1, col2 = st.columns(2)
    with col1:
        pdf_file_1 = st.file_uploader("上传第一个 PDF (包含 a,b,c)", type=["pdf"], key="c5_pdf1")
        pdf_file_2 = st.file_uploader("上传第二个 PDF (包含 i,j,k)", type=["pdf"], key="c5_pdf2")
    with col2:
        template_file = st.file_uploader("上传 Word 模板 (.docx)", type=["docx"], key="c5_word")

    if st.button("开始生成", key="btn_c5"):
        if pdf_file_1 and pdf_file_2 and template_file:
            try:
                with st.spinner("正在处理..."):
                    pdf_path_1 = save_uploaded_file(pdf_file_1)
                    pdf_path_2 = save_uploaded_file(pdf_file_2)
                    template_path = save_uploaded_file(template_file)
                    
                    output_filename = f"Generated_{os.path.splitext(template_file.name)[0]}.docx"
                    output_path = os.path.join(st.session_state.temp_dir, output_filename)

                    # PDF 1 处理
                    pdf_filename_1 = os.path.basename(pdf_path_1)
                    val_1 = extract_values_from_filename(pdf_filename_1)
                    
                    # PDF 2 处理
                    pdf_filename_2 = os.path.basename(pdf_path_2)
                    val_2 = extract_values_from_filename(pdf_filename_2)

                    if not val_1 or not val_2:
                        st.error("文件名格式错误，未找到足够的数值。")
                    else:
                        # 提取 d (PDF1)
                        def extract_row_values_local(pdf_path, page_num, keyword):
                            tables = camelot.read_pdf(pdf_path, pages=str(page_num), flavor='stream')
                            for table in tables:
                                df = table.df
                                for i, row in df.iterrows():
                                    if keyword in row.to_string():
                                        values = [val.replace(',', '') for val in re.findall(r"[\d,.]+", row.to_string())]
                                        return values
                            return []

                        d_values = extract_row_values_local(pdf_path_1, 3, "CIP2") or extract_row_values_local(pdf_path_1, 3, "CIM3")
                        d = d_values[3] if len(d_values) > 3 else "N/A"

                        # 提取 e, f, g, h (PDF1 Page 4)
                        num_rows_page_4 = 0
                        tables_page_4 = camelot.read_pdf(pdf_path_1, pages='4', flavor='stream')
                        for table in tables_page_4:
                            df_page_4 = table.df
                            num_rows_page_4 = df_page_4.shape[0]
                        
                        e = extract_table_value(pdf_path_1, 4, num_rows_page_4 - 8, 8)
                        f = extract_table_value(pdf_path_1, 4, num_rows_page_4 - 6, 8)
                        g = extract_table_value(pdf_path_1, 4, num_rows_page_4 - 4, 8)
                        h = extract_table_value(pdf_path_1, 4, num_rows_page_4 - 2, 8)

                        # 提取 l (PDF2)
                        l_values = extract_row_values_local(pdf_path_2, 3, "CIP2") or extract_row_values_local(pdf_path_2, 3, "CIM3")
                        l = l_values[3] if len(l_values) > 3 else "N/A"

                        # 提取 m, n, o, p (PDF2 Page 4)
                        num_rows_page_4_2 = 0
                        tables_page_4_2 = camelot.read_pdf(pdf_path_2, pages='4', flavor='stream')
                        for table in tables_page_4_2:
                            df_page_4_2 = table.df
                            num_rows_page_4_2 = df_page_4_2.shape[0]

                        m = extract_table_value(pdf_path_2, 4, num_rows_page_4_2 - 8, 8)
                        n = extract_table_value(pdf_path_2, 4, num_rows_page_4_2 - 6, 8)
                        o = extract_table_value(pdf_path_2, 4, num_rows_page_4_2 - 4, 8)
                        p = extract_table_value(pdf_path_2, 4, num_rows_page_4_2 - 2, 8)

                        # 合并数据
                        values = dict(zip("abc", val_1))
                        values.update(dict(zip("ijk", val_2)))
                        values.update({
                            "d": d, "e": e, "f": f, "g": g, "h": h,
                            "l": l, "m": m, "n": n, "o": o, "p": p
                        })

                        replace_values_in_word_template(template_path, output_path, values)

                        st.success("✅ 生成成功！")
                        with open(output_path, "rb") as file:
                            st.download_button(
                                label="下载生成的 Word 文档",
                                data=file,
                                file_name=output_filename,
                                mime="application/vnd.openxmlformats-officedocument.wordprocessingml.document"
                            )
            except Exception as e:
                st.error(f"发生错误: {e}")
        else:
            st.warning("请上传必要的文件。")
