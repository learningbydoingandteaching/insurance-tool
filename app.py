import streamlit as st
import os
import re
import camelot
import fitz  # PyMuPDF
from docx import Document
import copy
import tempfile
import shutil

# 设置页面配置
st.set_page_config(page_title="PDF 智能处理工具", layout="wide")

# ==========================================
# 1. 公共工具函数
# ==========================================

def extract_values_from_filename(filename):
    """从文件名提取前3个数字"""
    values = re.findall(r'\d+', filename)
    if len(values) >= 3:
        return values[:3]
    return None

def extract_table_value(pdf_path, page_num, row_num, col_num):
    """从指定页码、行、列提取表格数值 (通用)"""
    try:
        tables = camelot.read_pdf(pdf_path, pages=str(page_num), flavor='stream')
        for table in tables:
            df = table.df
            try:
                value = df.iat[int(row_num), int(col_num)]
                if value:
                    return value.replace(',', '')
            except IndexError:
                continue
        return "N/A"
    except Exception as e:
        return "N/A"

def extract_row_values(pdf_path, page_num, keyword):
    """搜索包含关键词的行，并提取该行所有数字"""
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
    """添加千位分隔符"""
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
    """计算 {{a+b}} 形式的表达式"""
    for key, value in values.items():
        val_str = str(value) if value != "N/A" else "0"
        expression = expression.replace(f"{{{key}}}", val_str)
    try:
        result = eval(expression)
        return add_thousand_separator(result)
    except Exception:
        return "N/A"

def replace_and_evaluate_in_run(run, values):
    """在 Word 的 Run 对象中执行替换"""
    full_text = run.text
    # 1. 直接替换 {key}
    for key, value in values.items():
        placeholder = f"{{{key}}}"
        full_text = full_text.replace(placeholder, str(value) if value is not None else "N/A")

    # 2. 计算 {{expression}}
    expressions = re.findall(r'\{\{[^\}]+\}\}', full_text)
    for expr in expressions:
        expr_clean = expr.strip("{}")
        result = evaluate_expression(expr_clean, values)
        full_text = full_text.replace(expr, str(result))

    run.text = full_text

def replace_and_evaluate_in_paragraph(paragraph, values):
    for run in paragraph.runs:
        replace_and_evaluate_in_run(run, values)

def replace_values_in_word_template(template_path, output_path, values):
    """遍历 Word 文档进行替换"""
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
# 2. 储蓄险专用函数
# ==========================================

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

def delete_specified_runs(doc, start_text, end_text):
    """删除指定范围内的文本"""
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
            try:
                p = paragraph._element
                p.getparent().remove(p)
                p._element = None
            except: pass

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
    if os.path.exists(output_path):
        doc = Document(output_path)
    else:
        doc = Document()
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

    for element in template_doc.element.body:
        doc.element.body.append(copy.deepcopy(element))
    doc.save(output_path)

# ==========================================
# 3. 核心修复逻辑：智能行搜索与求和
# ==========================================

def get_summed_value_by_age(df, target_age):
    """
    智能查找逻辑：
    1. 遍历表格每一行，寻找包含 target_age (如 "56") 的行。
    2. 找到行后，提取该行所有有效的数值。
    3. 取最后3个数值相加 (对应：保证值 + 红利A + 红利B)。
    """
    target_age_str = str(target_age)
    
    for index, row in df.iterrows():
        # 1. 将整行转为字符串列表
        row_str_list = [str(x).strip() for x in row.values]
        
        # 2. 检查这一行是否包含目标年龄 (通常在第1或第2列)
        # 我们检查前3列即可，防止误匹配到后面的金额
        found_age = False
        for cell in row_str_list[:3]:
            # 精确匹配 "56" 或者 "56.0"
            if cell == target_age_str or cell == f"{target_age_str}.0":
                found_age = True
                break
        
        if found_age:
            # 3. 提取该行所有的数值
            numbers = []
            for cell in row_str_list:
                # 去除逗号
                clean_cell = cell.replace(',', '').replace(' ', '')
                # 尝试转为浮点数
                try:
                    val = float(clean_cell)
                    numbers.append(val)
                except ValueError:
                    continue
            
            # 4. 逻辑推断：我们需要最后3个大数相加
            if len(numbers) >= 3:
                # 取最后三个数
                v1 = numbers[-1]
                v2 = numbers[-2]
                v3 = numbers[-3]
                
                # 简单的启发式规则：金额通常大于 200，防止把年龄加进去
                valid_values = [n for n in numbers if n > 200] 
                
                if len(valid_values) >= 3:
                    total = valid_values[-1] + valid_values[-2] + valid_values[-3]
                    return "{:,.0f}".format(total)
                else:
                    # 如果过滤后不足3个，直接加最后3个原始提取的数
                    total = v1 + v2 + v3
                    return "{:,.0f}".format(total)
            
            return "N/A (数据不足)"

    return "N/A"

# ==========================================
# 4. 业务处理流程
# ==========================================

def process_code1(pdf_file_path, new_pdf_file_path, template_path, output_path):
    # 1. 文件名提取
    pdf_filename = os.path.basename(pdf_file_path)
    filename_values = extract_values_from_filename_code1(pdf_filename)
    if not filename_values:
        return False, "PDF 文件名中未找到足够的数值 (需要至少6个数字)。"

    # 2. 基础数据提取 (g, h)
    doc = fitz.open(pdf_file_path)
    total_pages = len(doc)
    page_num_g_h = total_pages - 6
    g = extract_table_value(pdf_file_path, page_num_g_h, 11, 5)
    h = extract_table_value(pdf_file_path, page_num_g_h, 12, 5)

    # 3. 第6页复杂数据提取 (i, j, k, l, m) - 使用智能搜索逻辑
    tables_page_6 = camelot.read_pdf(pdf_file_path, pages='6', flavor='stream')
    i = j = k = l = m = "N/A"
    
    if len(tables_page_6) > 0:
        df_page_6 = tables_page_6[0].df
        # 智能搜索年龄行
        i = get_summed_value_by_age(df_page_6, 56)
        j = get_summed_value_by_age(df_page_6, 66)
        k = get_summed_value_by_age(df_page_6, 76)
        l = get_summed_value_by_age(df_page_6, 86)
        m = get_summed_value_by_age(df_page_6, 96)

    pdf_values = {"g": g, "h": h, "i": i, "j": j, "k": k, "l": l, "m": m}
    values = dict(zip("abcdef", filename_values))
    values.update(pdf_values)

    # 4. 生成文档 (无分阶段提取)
    if not new_pdf_file_path:
        remove_text_start = "在人生的重要阶段提取："
        remove_text_end = "不提取分红，在某年，把累积的本金"
        replace_values_in_word_template_with_delete(template_path, output_path, values, remove_text_start, remove_text_end)
        return True, "处理完成 (无分阶段提取)。"

    # 5. 处理分阶段提取 PDF
    new_pdf_filename = os.path.basename(new_pdf_file_path)
    n, o, p = extract_nop_from_filename(new_pdf_filename)
    if not n or not o or not p:
        return False, "新的 PDF 文件名中未找到 n, o, p 数值。"

    new_doc = fitz.open(new_pdf_file_path)
    total_new_pages = len(new_doc)
    page_num_q_r = total_new_pages - 6

    q = extract_table_value(new_pdf_file_path, page_num_q_r, 11, 5)
    r = extract_table_value(new_pdf_file_path, page_num_q_r, 12, 5)
    s_string = extract_table_value(new_pdf_file_path, page_num_q_r, 11, 0)
    s = extract_numeric_value_from_string(s_string)

    new_pdf_values = {"n": n, "o": o, "p": p, "q": q, "r": r, "s": s}
    values.update(new_pdf_values)

    replace_values_in_word_template(template_path, output_path, values)
    return True, "储蓄险处理完成！"

def process_code4(pdf_file_path, new_pdf_file_path, template_path, output_path):
    # 储蓄险添加逻辑
    pdf_filename = os.path.basename(pdf_file_path)
    filename_values = extract_values_from_filename_code1(pdf_filename)
    if not filename_values:
        return False, "PDF 文件名错误。"

    doc = fitz.open(pdf_file_path)
    total_pages = len(doc)
    page_num_g_h = total_pages - 6

    g = extract_table_value(pdf_file_path, page_num_g_h, 11, 5)
    h = extract_table_value(pdf_file_path, page_num_g_h, 12, 5)
    
    page_num_s = total_pages - 6
    s_string = extract_table_value(pdf_file_path, page_num_s, 11, 0)
    s = extract_numeric_value_from_string(s_string)

    # 第6页提取 - 使用智能搜索逻辑
    tables_page_6 = camelot.read_pdf(pdf_file_path, pages='6', flavor='stream')
    i = j = k = l = m = "N/A"
    if len(tables_page_6) > 0:
        df_page_6 = tables_page_6[0].df
        i = get_summed_value_by_age(df_page_6, 56)
        j = get_summed_value_by_age(df_page_6, 66)
        k = get_summed_value_by_age(df_page_6, 76)
        l = get_summed_value_by_age(df_page_6, 86)
        m = get_summed_value_by_age(df_page_6, 96)

    pdf_values = {"g": g, "h": h, "i": i, "j": j, "k": k, "l": l, "m": m, "s": s}
    values = dict(zip("abcdef", filename_values))
    values.update(pdf_values)

    if not new_pdf_file_path:
        remove_text_start = "在人生的重要阶段提取："
        remove_text_end = "不提取分红，在某年，把累积的本金"
        replace_values_in_word_template_append(template_path, output_path, values, remove_text_start, remove_text_end)
        return True, "储蓄险添加完成 (无分阶段)。"

    new_pdf_filename = os.path.basename(new_pdf_file_path)
    n, o, p = extract_nop_from_filename(new_pdf_filename)
    if not n or not o or not p:
        return False, "新PDF文件名错误。"

    new_doc = fitz.open(new_pdf_file_path)
    total_new_pages = len(new_doc)
    page_num_q_r = total_new_pages - 6
    q = extract_table_value(new_pdf_file_path, page_num_q_r, 11, 5)
    r = extract_table_value(new_pdf_file_path, page_num_q_r, 12, 5)

    values.update({"n": n, "o": o, "p": p, "q": q, "r": r})
    replace_values_in_word_template_append(template_path, output_path, values)
    return True, "储蓄险添加处理完成！"

def process_critical_illness(pdf_files, template_path, output_path, num_people):
    # 重疾险通用逻辑
    all_values = {}
    prefixes = [
        {"file_vars": ["a", "b", "c"], "pdf_vars": ["d", "e", "f", "g", "h"]},
        {"file_vars": ["a1", "b1", "c1"], "pdf_vars": ["d1", "e1", "f1", "g1", "h1"]},
        {"file_vars": ["a2", "b2", "c2"], "pdf_vars": ["d2", "e2", "f2", "g2", "h2"]},
        {"file_vars": ["a3", "b3", "c3"], "pdf_vars": ["d3", "e3", "f3", "g3", "h3"]},
    ]

    for idx in range(num_people):
        if idx >= len(pdf_files): break
        
        pdf_path = pdf_files[idx]
        pdf_filename = os.path.basename(pdf_path)
        
        filename_values = extract_values_from_filename(pdf_filename)
        if not filename_values:
            return False, f"第 {idx+1} 个PDF文件名中未找到足够数值。"
        
        d_values = extract_row_values(pdf_path, 3, "CIP2") or extract_row_values(pdf_path, 3, "CIM3")
        d = d_values[3] if len(d_values) > 3 else "N/A"

        num_rows_page_4 = 0
        tables_page_4 = camelot.read_pdf(pdf_path, pages='4', flavor='stream')
        e = f = g = h = "N/A"
        
        for table in tables_page_4:
            df_page_4 = table.df
            num_rows_page_4 = df_page_4.shape[0]
            if num_rows_page_4 > 8:
                e = extract_table_value(pdf_path, 4, num_rows_page_4 - 8, 8)
                f = extract_table_value(pdf_path, 4, num_rows_page_4 - 6, 8)
                g = extract_table_value(pdf_path, 4, num_rows_page_4 - 4, 8)
                h = extract_table_value(pdf_path, 4, num_rows_page_4 - 2, 8)

        prefix_config = prefixes[idx]
        all_values.update(dict(zip(prefix_config["file_vars"], filename_values)))
        all_values.update(dict(zip(prefix_config["pdf_vars"], [d, e, f, g, h])))

    replace_values_in_word_template(template_path, output_path, all_values)
    return True, f"{num_people}人重疾险处理完成！"

# ==========================================
# 5. Streamlit 界面主入口
# ==========================================

def save_uploaded_file(uploaded_file, temp_dir):
    """保存上传文件到临时目录，保持原文件名"""
    if uploaded_file is not None:
        file_path = os.path.join(temp_dir, uploaded_file.name)
        with open(file_path, "wb") as f:
            f.write(uploaded_file.getbuffer())
        return file_path
    return None

def main():
    st.title("📄 保险计划书自动化处理工具")
    
    st.markdown("### 📌 使用说明")
    st.info("""
    1. **文件名命名规范**非常重要，程序依赖文件名提取年龄、保额等信息。
    2. **储蓄险**：会自动计算退保价值（保证+红利），无需手动查找。
    """)

    # 侧边栏
    option = st.sidebar.radio(
        "选择操作类型",
        ["储蓄险", "储蓄险添加", "一人重疾险", "二人重疾险", "三人重疾险", "四人重疾险"]
    )

    # 文件上传区
    st.header("1. 上传文件")
    template_file = st.file_uploader("选择 Word 模板 (.docx)", type=["docx"])
    
    pdf_files = []
    new_pdf_file = None

    if option in ["储蓄险", "储蓄险添加"]:
        pdf_main = st.file_uploader("选择连续提取 PDF 文件 (必选)", type=["pdf"], key="main_pdf")
        if pdf_main: pdf_files.append(pdf_main)
        new_pdf_file = st.file_uploader("选择分阶段提取 PDF 文件 (可选)", type=["pdf"], key="sub_pdf")
        
    elif option == "一人重疾险":
        pdf = st.file_uploader("选择 PDF 文件", type=["pdf"], key="ci_1")
        if pdf: pdf_files.append(pdf)
        
    elif option == "二人重疾险":
        c1, c2 = st.columns(2)
        p1 = c1.file_uploader("PDF 1", type=["pdf"], key="ci_2_1")
        p2 = c2.file_uploader("PDF 2", type=["pdf"], key="ci_2_2")
        if p1 and p2: pdf_files = [p1, p2]
        
    elif option == "三人重疾险":
        c1, c2, c3 = st.columns(3)
        p1 = c1.file_uploader("PDF 1", type=["pdf"], key="ci_3_1")
        p2 = c2.file_uploader("PDF 2", type=["pdf"], key="ci_3_2")
        p3 = c3.file_uploader("PDF 3", type=["pdf"], key="ci_3_3")
        if p1 and p2 and p3: pdf_files = [p1, p2, p3]
        
    elif option == "四人重疾险":
        c1, c2 = st.columns(2)
        p1 = c1.file_uploader("PDF 1", type=["pdf"], key="ci_4_1")
        p2 = c2.file_uploader("PDF 2", type=["pdf"], key="ci_4_2")
        p3 = c1.file_uploader("PDF 3", type=["pdf"], key="ci_4_3")
        p4 = c2.file_uploader("PDF 4", type=["pdf"], key="ci_4_4")
        if p1 and p2 and p3 and p4: pdf_files = [p1, p2, p3, p4]

    # 处理按钮
    st.header("2. 开始处理")
    if st.button("运行处理程序", type="primary"):
        if not template_file:
            st.error("请上传 Word 模板文件！")
            return
        if not pdf_files:
            st.error("请上传所需的 PDF 文件！")
            return

        with tempfile.TemporaryDirectory() as temp_dir:
            try:
                # 保存文件到临时目录
                temp_template_path = save_uploaded_file(template_file, temp_dir)
                temp_output_path = os.path.join(temp_dir, f"Result_{template_file.name}")
                saved_pdf_paths = [save_uploaded_file(p, temp_dir) for p in pdf_files]
                saved_new_pdf_path = save_uploaded_file(new_pdf_file, temp_dir) if new_pdf_file else None

                success = False
                message = ""

                with st.spinner("正在解析 PDF 表格并生成文档，请稍候..."):
                    if option == "储蓄险":
                        success, message = process_code1(saved_pdf_paths[0], saved_new_pdf_path, temp_template_path, temp_output_path)
                    elif option == "储蓄险添加":
                        success, message = process_code4(saved_pdf_paths[0], saved_new_pdf_path, temp_template_path, temp_output_path)
                    elif option == "一人重疾险":
                        success, message = process_critical_illness(saved_pdf_paths, temp_template_path, temp_output_path, 1)
                    elif option == "二人重疾险":
                        success, message = process_critical_illness(saved_pdf_paths, temp_template_path, temp_output_path, 2)
                    elif option == "三人重疾险":
                        success, message = process_critical_illness(saved_pdf_paths, temp_template_path, temp_output_path, 3)
                    elif option == "四人重疾险":
                        success, message = process_critical_illness(saved_pdf_paths, temp_template_path, temp_output_path, 4)

                if success:
                    st.success(message)
                    with open(temp_output_path, "rb") as f:
                        st.download_button(
                            label="📥 下载生成的 Word 文档",
                            data=f,
                            file_name=f"Processed_{template_file.name}",
                            mime="application/vnd.openxmlformats-officedocument.wordprocessingml.document"
                        )
                else:
                    st.error(f"处理失败: {message}")

            except Exception as e:
                st.error(f"发生错误: {str(e)}")

if __name__ == "__main__":
    main()
