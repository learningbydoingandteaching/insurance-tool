import os
import re
import camelot
import fitz  # PyMuPDF
from docx import Document
import tkinter as tk
from tkinter import messagebox
from tkinter import filedialog
import copy  # 用于深度复制文档元素


# 公共函数部分

def extract_values_from_filename(filename):
    values = re.findall(r'\d+', filename)
    if len(values) >= 3:
        return values[:3]
    return None

def extract_table_value(pdf_path, page_num, row_num, col_num):
    tables = camelot.read_pdf(pdf_path, pages=str(page_num), flavor='stream')
    for table in tables:
        df = table.df
        try:
            value = df.iat[int(row_num), int(col_num)].replace(',', '')  # 去除逗号
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

# 储蓄险（code1）特有函数

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

# 储蓄险（code1）主函数

def main_code1(pdf_file, new_pdf_file, template_path, output_path):
    pdf_filename = os.path.basename(pdf_file)
    filename_values = extract_values_from_filename_code1(pdf_filename)
    if not filename_values:
        messagebox.showerror("错误", "PDF 文件名中未找到足够的数值。")
        return

    # 处理原始 PDF 文件
    doc = fitz.open(pdf_file)
    total_pages = len(doc)
    page_num_g_h = total_pages - 6

    g = extract_table_value(pdf_file, page_num_g_h, 11, 5)
    h = extract_table_value(pdf_file, page_num_g_h, 12, 5)

        # --- 修改开始：提取第6页（退保价值表）的 i, j, k, l, m 值 ---
    
    # 读取PDF第6页
    tables_page_6 = camelot.read_pdf(pdf_file, pages='6', flavor='stream')
    
    if len(tables_page_6) > 0:
        df_page_6 = tables_page_6[0].df
        num_rows_page_6 = df_page_6.shape[0]

        # 定义一个内部函数，通过倒数行号提取最后一列的数值
        def get_val_from_last_col(row_from_bottom):
            try:
                # row_from_bottom: 倒数第几行 (例如 1 代表最后一行)
                # pandas索引从0开始，所以行索引是 num_rows - row_from_bottom
                target_row_idx = num_rows_page_6 - row_from_bottom
                # -1 代表最后一列
                val = df_page_6.iat[target_row_idx, -2]
                return val.replace(',', '').replace(' ', '')
            except Exception as e:
                print(f"提取数值出错: {e}")
                return "N/A"

        # 根据 ANB 年龄倒推行号 (假设表尾是 ANB 101)
        i = get_val_from_last_col(10) # ANB 56 (倒数第10行) -> 240,547
        j = get_val_from_last_col(8)  # ANB 66 (倒数第8行)  -> 454,690
        k = get_val_from_last_col(6)  # ANB 76 (倒数第6行)  -> 853,672
        l = get_val_from_last_col(4)  # ANB 86 (倒数第4行)  -> 1,602,632
        m = get_val_from_last_col(2)  # ANB 96 (倒数第2行)  -> 3,008,582
    else:
        i = j = k = l = m = "N/A"

    # --- 修改结束 ---


    pdf_values = {
        "g": g,
        "h": h,
        "i": i,
        "j": j,
        "k": k,
        "l": l,
        "m": m
    }

    values = dict(zip("abcdef", filename_values))
    values.update(pdf_values)

    if not new_pdf_file:
        remove_text_start = "在人生的重要阶段提取："
        remove_text_end = "不提取分红，在某年，把累积的本金"
        replace_values_in_word_template_with_delete(template_path, output_path, values, remove_text_start, remove_text_end)
        messagebox.showinfo("完成", "处理完成，未使用分阶段提取的 PDF 文件。")
        return

    # 处理新的 PDF 文件
    new_pdf_filename = os.path.basename(new_pdf_file)
    n, o, p = extract_nop_from_filename(new_pdf_filename)
    if not n or not o or not p:
        messagebox.showerror("错误", "新的 PDF 文件名中未找到足够的数值用于 n, o, p。")
        return

    new_doc = fitz.open(new_pdf_file)
    total_new_pages = len(new_doc)
    page_num_q_r = total_new_pages - 6

    # 从新的 PDF 文件中提取 q、r、s
    q = extract_table_value(new_pdf_file, page_num_q_r, 11, 5)
    r = extract_table_value(new_pdf_file, page_num_q_r, 12, 5)

    s_string = extract_table_value(new_pdf_file, page_num_q_r, 11, 0)
    s = extract_numeric_value_from_string(s_string)

    new_pdf_values = {
        "n": n,
        "o": o,
        "p": p,
        "q": q,
        "r": r,
        "s": s
    }

    values.update(new_pdf_values)

    replace_values_in_word_template(template_path, output_path, values)
    messagebox.showinfo("完成", "储蓄险处理完成！")

# 储蓄险添加（code4）主函数

def main_code4(pdf_file, new_pdf_file, template_path, output_path):
    pdf_filename = os.path.basename(pdf_file)
    filename_values = extract_values_from_filename_code1(pdf_filename)
    if not filename_values:
        messagebox.showerror("错误", "PDF 文件名中未找到足够的数值。")
        return

    doc = fitz.open(pdf_file)
    total_pages = len(doc)
    page_num_g_h = total_pages - 6

    # 提取 g 和 h 的值
    g = extract_table_value(pdf_file, page_num_g_h, 11, 5)
    h = extract_table_value(pdf_file, page_num_g_h, 12, 5)

    # 提取 s 的值
    page_num_s = total_pages - 6
    s_string = extract_table_value(pdf_file, page_num_s, 11, 0)
    s = extract_numeric_value_from_string(s_string)

    # --- 新代码：直接从第6页提取 ---
    tables_page_6 = camelot.read_pdf(pdf_file, pages='6', flavor='stream')
    
    # 初始化默认值，防止读不到报错
    i = j = k = l = m = "N/A" 

    if len(tables_page_6) > 0:
        df_page_6 = tables_page_6[0].df
        num_rows_page_6 = df_page_6.shape[0]

        def get_val_from_last_col(row_from_bottom):
            try:
                target_row_idx = num_rows_page_6 - row_from_bottom
                val = df_page_6.iat[target_row_idx, -2] # -1 表示最后一列
                return val.replace(',', '').replace(' ', '')
            except Exception:
                return "N/A"

        # 根据倒数行数提取
        i = get_val_from_last_col(10) # ANB 56
        j = get_val_from_last_col(8)  # ANB 66
        k = get_val_from_last_col(6)  # ANB 76
        l = get_val_from_last_col(4)  # ANB 86
        m = get_val_from_last_col(2)  # ANB 96

    pdf_values = {
        "g": g,
        "h": h,
        "i": i,
        "j": j,
        "k": k,
        "l": l,
        "m": m,
        "s": s
    }

    values = dict(zip("abcdef", filename_values))
    values.update(pdf_values)

    if not new_pdf_file:
        remove_text_start = "在人生的重要阶段提取："
        remove_text_end = "不提取分红，在某年，把累积的本金"
        replace_values_in_word_template_append(template_path, output_path, values, remove_text_start, remove_text_end)
        messagebox.showinfo("完成", "储蓄险添加处理完成，未使用分阶段提取的 PDF 文件。")
        return

    new_pdf_filename = os.path.basename(new_pdf_file)
    n, o, p = extract_nop_from_filename(new_pdf_filename)
    if not n or not o or not p:
        messagebox.showerror("错误", "新的 PDF 文件名中未找到足够的数值用于 n, o, p。")
        return

    new_doc = fitz.open(new_pdf_file)
    total_new_pages = len(new_doc)
    page_num_q_r = total_new_pages - 6

    # 提取 q 和 r 的值
    q = extract_table_value(new_pdf_file, page_num_q_r, 11, 5)
    r = extract_table_value(new_pdf_file, page_num_q_r, 12, 5)

    new_pdf_values = {
        "n": n,
        "o": o,
        "p": p,
        "q": q,
        "r": r
    }

    values.update(new_pdf_values)

    replace_values_in_word_template_append(template_path, output_path, values)
    messagebox.showinfo("完成", "储蓄险添加处理完成！")

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

# 一人重疾险（code2）主函数

def main_code2(pdf_file, template_path, output_path):
    pdf_filename = os.path.basename(pdf_file)
    filename_values = extract_values_from_filename(pdf_filename)
    if not filename_values:
        messagebox.showerror("错误", "PDF 文件名中未找到足够的数值。")
        return

    doc = fitz.open(pdf_file)
    page_num = 3

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

    pdf_values = {
        "d": d,
        "e": e,
        "f": f,
        "g": g,
        "h": h,
    }

    values = dict(zip("abc", filename_values))
    values.update(pdf_values)

    replace_values_in_word_template(template_path, output_path, values)
    messagebox.showinfo("完成", "重疾险处理完成！")

# 二人重疾险（code5）主函数

def main_code5(pdf_file_1, pdf_file_2, template_path, output_path):
    # 提取第一个PDF
    pdf_filename_1 = os.path.basename(pdf_file_1)
    if not pdf_filename_1.endswith('.pdf'):
        messagebox.showerror("错误", "第一个文件不是PDF文件。")
        return

    filename_values_1 = extract_values_from_filename(pdf_filename_1)
    if not filename_values_1:
        messagebox.showerror("错误", "第一个PDF文件名中未找到足够的数值。")
        return

    # 提取第二个PDF
    pdf_filename_2 = os.path.basename(pdf_file_2)
    if not pdf_filename_2.endswith('.pdf'):
        messagebox.showerror("错误", "第二个文件不是PDF文件。")
        return

    filename_values_2 = extract_values_from_filename(pdf_filename_2)
    if not filename_values_2:
        messagebox.showerror("错误", "第二个PDF文件名中未找到足够的数值。")
        return

    # 后续处理
    # 使用pdf_file_1和pdf_file_2作为PDF文件的路径
    # 不需要再组合路径或列出目录

    # 处理第一个PDF
    d_values = extract_row_values(pdf_file_1, 3, "CIP2") or extract_row_values(pdf_file_1, 3, "CIM3")
    d = d_values[3] if len(d_values) > 3 else "N/A"

    num_rows_page_4 = 0
    tables_page_4 = camelot.read_pdf(pdf_file_1, pages='4', flavor='stream')
    for table in tables_page_4:
        df_page_4 = table.df
        num_rows_page_4 = df_page_4.shape[0]

    e = extract_table_value(pdf_file_1, 4, num_rows_page_4 - 8, 8)
    f = extract_table_value(pdf_file_1, 4, num_rows_page_4 - 6, 8)
    g = extract_table_value(pdf_file_1, 4, num_rows_page_4 - 4, 8)
    h = extract_table_value(pdf_file_1, 4, num_rows_page_4 - 2, 8)

    pdf_values_1 = {
        "d": d,
        "e": e,
        "f": f,
        "g": g,
        "h": h,
    }

    # 处理第二个PDF
    d1_values = extract_row_values(pdf_file_2, 3, "CIP2") or extract_row_values(pdf_file_2, 3, "CIM3")
    d1 = d1_values[3] if len(d1_values) > 3 else "N/A"

    num_rows_page_4_2 = 0
    tables_page_4_2 = camelot.read_pdf(pdf_file_2, pages='4', flavor='stream')
    for table in tables_page_4_2:
        df_page_4_2 = table.df
        num_rows_page_4_2 = df_page_4_2.shape[0]

    e1 = extract_table_value(pdf_file_2, 4, num_rows_page_4_2 - 8, 8)
    f1 = extract_table_value(pdf_file_2, 4, num_rows_page_4_2 - 6, 8)
    g1 = extract_table_value(pdf_file_2, 4, num_rows_page_4_2 - 4, 8)
    h1 = extract_table_value(pdf_file_2, 4, num_rows_page_4_2 - 2, 8)

    pdf_values_2 = {
        "d1": d1,
        "e1": e1,
        "f1": f1,
        "g1": g1,
        "h1": h1,
    }

    # 合并所有值并进行替换
    values = dict(zip(["a", "b", "c"], filename_values_1))
    values.update(pdf_values_1)
    values.update(dict(zip(["a1", "b1", "c1"], filename_values_2)))
    values.update(pdf_values_2)

    replace_values_in_word_template(template_path, output_path, values)
    messagebox.showinfo("完成", "二人重疾险处理完成！")

# 三人重疾险（code6）主函数

def main_code6(pdf_file_1, pdf_file_2, pdf_file_3, template_path, output_path):
    # 提取第一个PDF
    pdf_filename_1 = os.path.basename(pdf_file_1)
    if not pdf_filename_1.endswith('.pdf'):
        messagebox.showerror("错误", "第一个文件不是PDF文件。")
        return

    filename_values_1 = extract_values_from_filename(pdf_filename_1)
    if not filename_values_1:
        messagebox.showerror("错误", "第一个PDF文件名中未找到足够的数值。")
        return

    # 提取第二个PDF
    pdf_filename_2 = os.path.basename(pdf_file_2)
    if not pdf_filename_2.endswith('.pdf'):
        messagebox.showerror("错误", "第二个文件不是PDF文件。")
        return


    filename_values_2 = extract_values_from_filename(pdf_filename_2)
    if not filename_values_2:
        messagebox.showerror("错误", "第二个PDF文件名中未找到足够的数值。")
        return

    # 提取第三个PDF
    pdf_filename_3 = os.path.basename(pdf_file_3)
    if not pdf_filename_3.endswith('.pdf'):
        messagebox.showerror("错误", "第三个文件不是PDF文件。")
        return

    filename_values_3 = extract_values_from_filename(pdf_filename_3)
    if not filename_values_3:
        messagebox.showerror("错误", "第三个PDF文件名中未找到足够的数值。")
        return

    # 处理第一个PDF
    d_values = extract_row_values(pdf_file_1, 3, "CIP2") or extract_row_values(pdf_file_1, 3, "CIM3")
    d = d_values[3] if len(d_values) > 3 else "N/A"

    num_rows_page_4 = 0
    tables_page_4 = camelot.read_pdf(pdf_file_1, pages='4', flavor='stream')
    for table in tables_page_4:
        df_page_4 = table.df
        num_rows_page_4 = df_page_4.shape[0]

    e = extract_table_value(pdf_file_1, 4, num_rows_page_4 - 8, 8)
    f = extract_table_value(pdf_file_1, 4, num_rows_page_4 - 6, 8)
    g = extract_table_value(pdf_file_1, 4, num_rows_page_4 - 4, 8)
    h = extract_table_value(pdf_file_1, 4, num_rows_page_4 - 2, 8)

    pdf_values_1 = {
        "d": d,
        "e": e,
        "f": f,
        "g": g,
        "h": h,
    }

    # 处理第二个PDF
    d1_values = extract_row_values(pdf_file_2, 3, "CIP2") or extract_row_values(pdf_file_2, 3, "CIM3")
    d1 = d1_values[3] if len(d1_values) > 3 else "N/A"

    num_rows_page_4_2 = 0
    tables_page_4_2 = camelot.read_pdf(pdf_file_2, pages='4', flavor='stream')
    for table in tables_page_4_2:
        df_page_4_2 = table.df
        num_rows_page_4_2 = df_page_4_2.shape[0]

    e1 = extract_table_value(pdf_file_2, 4, num_rows_page_4_2 - 8, 8)
    f1 = extract_table_value(pdf_file_2, 4, num_rows_page_4_2 - 6, 8)
    g1 = extract_table_value(pdf_file_2, 4, num_rows_page_4_2 - 4, 8)
    h1 = extract_table_value(pdf_file_2, 4, num_rows_page_4_2 - 2, 8)

    pdf_values_2 = {
        "d1": d1,
        "e1": e1,
        "f1": f1,
        "g1": g1,
        "h1": h1,
    }

    # 处理第三个PDF
    d2_values = extract_row_values(pdf_file_3, 3, "CIP2") or extract_row_values(pdf_file_3, 3, "CIM3")
    d2 = d2_values[3] if len(d2_values) > 3 else "N/A"

    num_rows_page_4_3 = 0
    tables_page_4_3 = camelot.read_pdf(pdf_file_3, pages='4', flavor='stream')
    for table in tables_page_4_3:
        df_page_4_3 = table.df
        num_rows_page_4_3 = df_page_4_3.shape[0]

    e2 = extract_table_value(pdf_file_3, 4, num_rows_page_4_3 - 8, 8)
    f2 = extract_table_value(pdf_file_3, 4, num_rows_page_4_3 - 6, 8)
    g2 = extract_table_value(pdf_file_3, 4, num_rows_page_4_3 - 4, 8)
    h2 = extract_table_value(pdf_file_3, 4, num_rows_page_4_3 - 2, 8)

    pdf_values_3 = {
        "d2": d2,
        "e2": e2,
        "f2": f2,
        "g2": g2,
        "h2": h2
    }

    # 合并所有值并进行替换
    values = dict(zip(["a", "b", "c"], filename_values_1))
    values.update(pdf_values_1)
    values.update(dict(zip(["a1", "b1", "c1"], filename_values_2)))
    values.update(pdf_values_2)
    values.update(dict(zip(["a2", "b2", "c2"], filename_values_3)))
    values.update(pdf_values_3)

    replace_values_in_word_template(template_path, output_path, values)
    messagebox.showinfo("完成", "三人重疾险处理完成！")

def main_code7(pdf_file_1, pdf_file_2, pdf_file_3, pdf_file_4, template_path, output_path):
    # 提取第一个PDF
    pdf_filename_1 = os.path.basename(pdf_file_1)
    if not pdf_filename_1.endswith('.pdf'):
        messagebox.showerror("错误", "第一个文件不是PDF文件。")
        return

    filename_values_1 = extract_values_from_filename(pdf_filename_1)
    if not filename_values_1:
        messagebox.showerror("错误", "第一个PDF文件名中未找到足够的数值。")
        return

    # 提取第二个PDF
    pdf_filename_2 = os.path.basename(pdf_file_2)
    if not pdf_filename_2.endswith('.pdf'):
        messagebox.showerror("错误", "第二个文件不是PDF文件。")
        return

    filename_values_2 = extract_values_from_filename(pdf_filename_2)
    if not filename_values_2:
        messagebox.showerror("错误", "第二个PDF文件名中未找到足够的数值。")
        return

    # 提取第三个PDF
    pdf_filename_3 = os.path.basename(pdf_file_3)
    if not pdf_filename_3.endswith('.pdf'):
        messagebox.showerror("错误", "第三个文件不是PDF文件。")
        return

    filename_values_3 = extract_values_from_filename(pdf_filename_3)
    if not filename_values_3:
        messagebox.showerror("错误", "第三个PDF文件名中未找到足够的数值。")
        return

    # 提取第四个PDF
    pdf_filename_4 = os.path.basename(pdf_file_4)
    if not pdf_filename_4.endswith('.pdf'):
        messagebox.showerror("错误", "第四个文件不是PDF文件。")
        return

    filename_values_4 = extract_values_from_filename(pdf_filename_4)
    if not filename_values_4:
        messagebox.showerror("错误", "第四个PDF文件名中未找到足够的数值。")
        return

    # 处理第一个PDF
    d_values = extract_row_values(pdf_file_1, 3, "CIP2") or extract_row_values(pdf_file_1, 3, "CIM3")
    d = d_values[3] if len(d_values) > 3 else "N/A"

    num_rows_page_4 = 0
    tables_page_4 = camelot.read_pdf(pdf_file_1, pages='4', flavor='stream')
    for table in tables_page_4:
        df_page_4 = table.df
        num_rows_page_4 = df_page_4.shape[0]

    e = extract_table_value(pdf_file_1, 4, num_rows_page_4 - 8, 8)
    f = extract_table_value(pdf_file_1, 4, num_rows_page_4 - 6, 8)
    g = extract_table_value(pdf_file_1, 4, num_rows_page_4 - 4, 8)
    h = extract_table_value(pdf_file_1, 4, num_rows_page_4 - 2, 8)

    pdf_values_1 = {
        "d": d,
        "e": e,
        "f": f,
        "g": g,
        "h": h,
    }

    # 处理第二个PDF
    d1_values = extract_row_values(pdf_file_2, 3, "CIP2") or extract_row_values(pdf_file_2, 3, "CIM3")
    d1 = d1_values[3] if len(d1_values) > 3 else "N/A"

    num_rows_page_4_2 = 0
    tables_page_4_2 = camelot.read_pdf(pdf_file_2, pages='4', flavor='stream')
    for table in tables_page_4_2:
        df_page_4_2 = table.df
        num_rows_page_4_2 = df_page_4_2.shape[0]

    e1 = extract_table_value(pdf_file_2, 4, num_rows_page_4_2 - 8, 8)
    f1 = extract_table_value(pdf_file_2, 4, num_rows_page_4_2 - 6, 8)
    g1 = extract_table_value(pdf_file_2, 4, num_rows_page_4_2 - 4, 8)
    h1 = extract_table_value(pdf_file_2, 4, num_rows_page_4_2 - 2, 8)

    pdf_values_2 = {
        "d1": d1,
        "e1": e1,
        "f1": f1,
        "g1": g1,
        "h1": h1,
    }

    # 处理第三个PDF
    d2_values = extract_row_values(pdf_file_3, 3, "CIP2") or extract_row_values(pdf_file_3, 3, "CIM3")
    d2 = d2_values[3] if len(d2_values) > 3 else "N/A"

    num_rows_page_4_3 = 0
    tables_page_4_3 = camelot.read_pdf(pdf_file_3, pages='4', flavor='stream')
    for table in tables_page_4_3:
        df_page_4_3 = table.df
        num_rows_page_4_3 = df_page_4_3.shape[0]

    e2 = extract_table_value(pdf_file_3, 4, num_rows_page_4_3 - 8, 8)
    f2 = extract_table_value(pdf_file_3, 4, num_rows_page_4_3 - 6, 8)
    g2 = extract_table_value(pdf_file_3, 4, num_rows_page_4_3 - 4, 8)
    h2 = extract_table_value(pdf_file_3, 4, num_rows_page_4_3 - 2, 8)

    pdf_values_3 = {
        "d2": d2,
        "e2": e2,
        "f2": f2,
        "g2": g2,
        "h2": h2
    }

    # 处理第四个PDF
    d3_values = extract_row_values(pdf_file_4, 3, "CIP2") or extract_row_values(pdf_file_4, 3, "CIM3")
    d3 = d3_values[3] if len(d3_values) > 3 else "N/A"

    num_rows_page_4_4 = 0
    tables_page_4_4 = camelot.read_pdf(pdf_file_4, pages='4', flavor='stream')
    for table in tables_page_4_4:
        df_page_4_4 = table.df
        num_rows_page_4_4 = df_page_4_4.shape[0]

    e3 = extract_table_value(pdf_file_4, 4, num_rows_page_4_4 - 8, 8)
    f3 = extract_table_value(pdf_file_4, 4, num_rows_page_4_4 - 6, 8)
    g3 = extract_table_value(pdf_file_4, 4, num_rows_page_4_4 - 4, 8)
    h3 = extract_table_value(pdf_file_4, 4, num_rows_page_4_4 - 2, 8)

    pdf_values_4 = {
        "d3": d3,
        "e3": e3,
        "f3": f3,
        "g3": g3,
        "h3": h3
    }

    # 合并所有值并进行替换
    values = dict(zip(["a", "b", "c"], filename_values_1))
    values.update(pdf_values_1)
    values.update(dict(zip(["a1", "b1", "c1"], filename_values_2)))
    values.update(pdf_values_2)
    values.update(dict(zip(["a2", "b2", "c2"], filename_values_3)))
    values.update(pdf_values_3)
    values.update(dict(zip(["a3", "b3", "c3"], filename_values_4)))
    values.update(pdf_values_4)

    replace_values_in_word_template(template_path, output_path, values)
    messagebox.showinfo("完成", "四人重疾险处理完成！")


# 图形界面部分

def main():
    root = tk.Tk()
    root.title("PDF处理工具")

    # 选项类型
    option_var = tk.StringVar(value="储蓄险")

    def update_fields(*args):
        option = option_var.get()
        if option in ["储蓄险", "储蓄险添加"]:
            pdf_file_label.config(text="选择连续提取PDF文件：")
            new_pdf_file_label.config(state="normal")
            new_pdf_file_entry.config(state="normal")
            new_pdf_file_button.config(state="normal")
            pdf_file_multi_frame.pack_forget()
            pdf_file_frame.pack(fill="x")
        elif option in ["二人重疾险", "三人重疾险", "四人重疾险"]:
            pdf_file_label.config(text="选择PDF文件：")
            new_pdf_file_label.config(state="disabled")
            new_pdf_file_entry.config(state="disabled")
            new_pdf_file_button.config(state="disabled")
            pdf_file_frame.pack_forget()
            pdf_file_multi_frame.pack(fill="x")
            update_pdf_file_fields(option)
        else:
            pdf_file_label.config(text="选择PDF文件：")
            new_pdf_file_label.config(state="disabled")
            new_pdf_file_entry.config(state="disabled")
            new_pdf_file_button.config(state="disabled")
            pdf_file_multi_frame.pack_forget()
            pdf_file_frame.pack(fill="x")
    def update_pdf_file_fields(option):
                num_files = {"二人重疾险": 2, "三人重疾险": 3, "四人重疾险": 4}
                n = num_files.get(option, 1)
                for i in range(4):
                    if i < n:
                        file_labels[i].pack(anchor="w", padx=5, pady=2)
                        file_entries[i].pack(fill="x", padx=5)
                        file_buttons[i].pack(padx=5, pady=2)
                    else:
                        file_labels[i].pack_forget()
                        file_entries[i].pack_forget()
                        file_buttons[i].pack_forget()

    def update_pdf_folder_fields(option):
        num_folders = {"二人重疾险": 2, "三人重疾险": 3, "四人重疾险": 4}
        n = num_folders.get(option, 1)
        for i in range(4):
            if i < n:
                folder_labels[i].grid(row=i, column=0, sticky="e", padx=5, pady=5)
                folder_entries[i].grid(row=i, column=1, padx=5, pady=5)
                folder_buttons[i].grid(row=i, column=2, padx=5, pady=5)
            else:
                folder_labels[i].grid_remove()
                folder_entries[i].grid_remove()
                folder_buttons[i].grid_remove()

    option_var.trace("w", update_fields)

    # 选项选择
    option_frame = tk.LabelFrame(root, text="选择操作类型")
    option_frame.pack(fill="x", padx=10, pady=5)

    options = ["储蓄险", "储蓄险添加", "一人重疾险", "二人重疾险", "三人重疾险", "四人重疾险"]
    for opt in options:
        tk.Radiobutton(option_frame, text=opt, variable=option_var, value=opt).pack(side="left", padx=5, pady=5)

        # 添加两行文本
    info_label_1 = tk.Label(root, text="PDF命名示例")
    info_label_1.pack(fill="x", padx=10, pady=2)

    info_label_2 = tk.Label(root, text="连续提取：4岁人士存20000美金存5年_19到85岁提取12000", anchor="w")
    info_label_2.pack(fill="x", padx=30, pady=2)

    info_label_3 = tk.Label(root,
                            text="分阶段提取：6岁人士存10000美金存5年_19到22岁提取8000_31岁提取20000_61到85岁提取31000",
                            anchor="w")
    info_label_3.pack(fill="x", padx=30, pady=2)

    info_label_4 = tk.Label(root,
                            text="重疾险：4岁男孩_10万美金起始保额_新加倍保20年供", anchor="w")
    info_label_4.pack(fill="x", padx=30, pady=2)

    # 输入文件选择
    input_frame = tk.LabelFrame(root, text="文件选择")
    input_frame.pack(fill="x", padx=10, pady=5)

    # PDF 文件框架
    pdf_file_frame = tk.Frame(input_frame)
    pdf_file_frame.pack(fill="x")

    # PDF 文件
    pdf_file_label = tk.Label(pdf_file_frame, text="选择PDF文件：")
    pdf_file_label.grid(row=0, column=0, sticky="e", padx=5, pady=5)
    pdf_file_entry = tk.Entry(pdf_file_frame, width=50)
    pdf_file_entry.grid(row=0, column=1, padx=5, pady=5)
    pdf_file_button = tk.Button(pdf_file_frame, text="浏览", command=lambda: select_file(pdf_file_entry, [("PDF文件", "*.pdf")]))
    pdf_file_button.grid(row=0, column=2, padx=5, pady=5)

    # 新的 PDF 文件（仅储蓄险和储蓄险添加）
    new_pdf_file_label = tk.Label(pdf_file_frame, text="选择分阶段提取PDF文件：")
    new_pdf_file_label.grid(row=1, column=0, sticky="e", padx=5, pady=5)
    new_pdf_file_entry = tk.Entry(pdf_file_frame, width=50)
    new_pdf_file_entry.grid(row=1, column=1, padx=5, pady=5)
    new_pdf_file_button = tk.Button(pdf_file_frame, text="浏览", command=lambda: select_file(new_pdf_file_entry, [("PDF文件", "*.pdf")]))
    new_pdf_file_button.grid(row=1, column=2, padx=5, pady=5)

    # PDF 文件框架（用于选择PDF文件）
    pdf_file_multi_frame = tk.Frame(input_frame)

    file_labels = []
    file_entries = []
    file_buttons = []

    for i in range(4):
        label = tk.Label(pdf_file_multi_frame, text=f"选择PDF文件{i + 1}：")
        entry = tk.Entry(pdf_file_multi_frame, width=50)
        button = tk.Button(pdf_file_multi_frame, text="浏览",
                           command=lambda e=entry: select_file(e, [("PDF文件", "*.pdf")]))
        file_labels.append(label)
        file_entries.append(entry)
        file_buttons.append(button)

    # Word 模板文件
    template_file_label = tk.Label(input_frame, text="选择Word模板文件：")
    template_file_label.pack(anchor="w", padx=5, pady=5)
    template_file_entry = tk.Entry(input_frame, width=50)
    template_file_entry.pack(fill="x", padx=5, pady=5)
    template_file_button = tk.Button(input_frame, text="浏览", command=lambda: select_file(template_file_entry, [("Word文件", "*.docx")]))
    template_file_button.pack(padx=5, pady=5)

    # 输出 Word 文件
    output_file_label = tk.Label(input_frame, text="保存输出的Word文件：")
    output_file_label.pack(anchor="w", padx=5, pady=5)
    output_file_entry = tk.Entry(input_frame, width=50)
    output_file_entry.pack(fill="x", padx=5, pady=5)
    output_file_button = tk.Button(input_frame, text="浏览", command=lambda: save_file(output_file_entry, [("Word文件", "*.docx")]))
    output_file_button.pack(padx=5, pady=5)

    # 更新字段状态
    update_fields()

    # 开始按钮
    start_button = tk.Button(root, text="开始处理", command=lambda: run_processing(
    option_var.get(),
    pdf_file_entry.get(),
    new_pdf_file_entry.get(),
    [e.get() for e in file_entries],
    template_file_entry.get(),
    output_file_entry.get()
))

    start_button.pack(pady=10)

    root.mainloop()

def select_file(entry_widget, filetypes):
    filepath = filedialog.askopenfilename(filetypes=filetypes)
    if filepath:
        entry_widget.delete(0, tk.END)
        entry_widget.insert(0, filepath)


def save_file(entry_widget, filetypes):
    filepath = filedialog.asksaveasfilename(defaultextension=filetypes[0][1], filetypes=filetypes)
    if filepath:
        entry_widget.delete(0, tk.END)
        entry_widget.insert(0, filepath)

def run_processing(option, pdf_file, new_pdf_file, pdf_files, template_file, output_file):
    if not template_file or not output_file:
        messagebox.showwarning("警告", "请提供所有必要的文件路径。")
        return

    if option == "储蓄险":
        if not pdf_file:
            messagebox.showwarning("警告", "请提供PDF文件路径。")
            return
        main_code1(pdf_file, new_pdf_file, template_file, output_file)
    elif option == "储蓄险添加":
        if not pdf_file:
            messagebox.showwarning("警告", "请提供PDF文件路径。")
            return
        main_code4(pdf_file, new_pdf_file, template_file, output_file)
    elif option == "一人重疾险":
        if not pdf_file:
            messagebox.showwarning("警告", "请提供PDF文件路径。")
            return
        main_code2(pdf_file, template_file, output_file)
    elif option == "二人重疾险":
        if len(pdf_files) < 2 or not all(pdf_files[:2]):
            messagebox.showwarning("警告", "请提供两个PDF文件路径。")
            return
        main_code5(pdf_files[0], pdf_files[1], template_file, output_file)
    elif option == "三人重疾险":
        if len(pdf_files) < 3 or not all(pdf_files[:3]):
            messagebox.showwarning("警告", "请提供三个PDF文件路径。")
            return
        main_code6(pdf_files[0], pdf_files[1], pdf_files[2], template_file, output_file)
    elif option == "四人重疾险":
        if len(pdf_files) < 4 or not all(pdf_files[:4]):
            messagebox.showwarning("警告", "请提供四个PDF文件路径。")
            return
        main_code7(pdf_files[0], pdf_files[1], pdf_files[2], pdf_files[3], template_file, output_file)


if __name__ == "__main__":
    main()
