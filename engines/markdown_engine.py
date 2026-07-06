"""Markdown 引擎：Markdown ↔ docx 转换（含数学公式）。"""

import os
import re

import mistune
from docx import Document
from docx.shared import Pt, Inches, RGBColor
from docx.enum.text import WD_ALIGN_PARAGRAPH
from latex2mathml.converter import convert as latex_to_mathml


def _render_md_ast_to_docx(doc, ast):
    """遍历 mistune AST，将节点追加到 python-docx Document。"""
    for node in ast:
        node_type = node.get("type", "")
        children = node.get("children", [])

        if node_type == "heading":
            level = node.get("attrs", {}).get("level", 1)
            para = doc.add_heading(level=min(level, 9))
            _render_children(para, children)
        elif node_type == "paragraph":
            para = doc.add_paragraph()
            _render_children(para, children)
        elif node_type == "block_code":
            code_text = node.get("raw", "")
            para = doc.add_paragraph()
            run = para.add_run(code_text)
            run.font.name = "Consolas"
            run.font.size = Pt(9)
        elif node_type == "block_quote":
            para = doc.add_paragraph()
            para.style = doc.styles["Quote"]
            _render_children(para, children)
        elif node_type == "list":
            ordered = node.get("attrs", {}).get("ordered", False)
            _render_list(doc, children, ordered)
        elif node_type == "thematic_break":
            doc.add_paragraph("─" * 60)
        elif node_type == "table":
            _render_table(doc, node)
        elif node_type == "block_html":
            pass
        elif node_type == "block_math":
            math_text = node.get("raw", "")
            _add_math(doc, math_text)


def _render_children(para, children):
    """将 mistune inline 子节点渲染到段落。"""
    for child in children:
        ctype = child.get("type", "")
        text = child.get("raw", "") or child.get("text", "")

        if ctype == "text":
            para.add_run(text)
        elif ctype == "strong":
            run = para.add_run(_collect_text(child))
            run.bold = True
        elif ctype == "emphasis":
            run = para.add_run(_collect_text(child))
            run.italic = True
        elif ctype == "codespan":
            run = para.add_run(child.get("raw", ""))
            run.font.name = "Consolas"
            run.font.size = Pt(9)
        elif ctype == "link":
            text = _collect_text(child)
            url = child.get("attrs", {}).get("url", "")
            para.add_run(f"{text} ({url})")
        elif ctype == "image":
            alt = child.get("attrs", {}).get("alt", "image")
            src = child.get("attrs", {}).get("url", "")
            para.add_run(f"\n[图片: {alt}]\n")
        elif ctype == "linebreak":
            para.add_run("\n")
        elif ctype == "inline_math":
            math_text = child.get("raw", "")
            _add_math_to_paragraph(para, math_text)
        elif ctype == "softbreak":
            para.add_run(" ")


def _collect_text(node):
    """递归收集节点下所有纯文本。"""
    texts = []
    if "children" in node:
        for child in node["children"]:
            if child.get("type") == "text":
                texts.append(child.get("raw", ""))
            elif "children" in child:
                texts.append(_collect_text(child))
    return "".join(texts)


def _render_list(doc, items, ordered):
    """渲染列表（有序/无序）。"""
    for i, item in enumerate(items):
        prefix = f"{i + 1}. " if ordered else "• "
        para = doc.add_paragraph()
        run = para.add_run(prefix)
        _render_children(para, item.get("children", []))


def _render_table(doc, node):
    """渲染 GFM 表格。"""
    header = node.get("children", [])[0] if node.get("children") else None
    body = node.get("children", [])[1] if len(node.get("children", [])) > 1 else None
    rows_data = []

    if header:
        rows_data.append([
            _collect_text(cell) for cell in header.get("children", [])
        ])
    if body:
        for row_node in body.get("children", []):
            rows_data.append([
                _collect_text(cell) for cell in row_node.get("children", [])
            ])

    if not rows_data:
        return
    table = doc.add_table(rows=len(rows_data), cols=len(rows_data[0]))
    table.style = "Table Grid"
    for i, row_data in enumerate(rows_data):
        for j, cell_text in enumerate(row_data):
            table.cell(i, j).text = cell_text


def _add_math(doc, math_tex):
    """在 docx 中添加块级数学公式。"""
    try:
        mathml = latex_to_mathml(math_tex)
        para = doc.add_paragraph()
        para.alignment = WD_ALIGN_PARAGRAPH.CENTER
        run = para.add_run(f"[公式: {math_tex}]")
        run.font.size = Pt(10)
        run.font.color.rgb = RGBColor(100, 100, 100)
    except Exception:
        para = doc.add_paragraph()
        para.alignment = WD_ALIGN_PARAGRAPH.CENTER
        run = para.add_run(f"$${math_tex}$$")
        run.font.size = Pt(10)


def _add_math_to_paragraph(para, math_tex):
    """在段落中添加行内数学公式。"""
    try:
        mathml = latex_to_mathml(math_tex)
        run = para.add_run(f"[{math_tex}]")
        run.font.color.rgb = RGBColor(100, 100, 100)
    except Exception:
        para.add_run(f"${math_tex}$")


def md_to_docx(input_path, output_path=None):
    """将 Markdown 文件转为 docx。返回 (成功, 消息)。"""
    if output_path is None:
        output_path = os.path.splitext(input_path)[0] + ".docx"

    try:
        with open(input_path, "r", encoding="utf-8") as f:
            md_text = f.read()

        renderer = mistune.create_markdown(renderer=None)
        ast = renderer(md_text)

        doc = Document()
        if isinstance(ast, list) and len(ast) > 0:
            _render_md_ast_to_docx(doc, ast)
        elif isinstance(ast, dict):
            _render_md_ast_to_docx(doc, [ast])
        else:
            for line in md_text.strip().split("\n"):
                doc.add_paragraph(line)

        doc.save(output_path)
        return True, f"✓ 完成：{os.path.basename(output_path)}"
    except Exception as e:
        return False, f"✗ 转换失败：{str(e)}"


def docx_to_md(input_path, output_path=None):
    """将 docx 文件转为 Markdown。返回 (成功, 消息)。"""
    if output_path is None:
        output_path = os.path.splitext(input_path)[0] + ".md"

    try:
        doc = Document(input_path)
        lines = []

        for para in doc.paragraphs:
            text = para.text.strip()
            if not text:
                lines.append("")
                continue

            style = para.style.name.lower() if para.style else ""

            if style.startswith("heading"):
                try:
                    level = int(style.split()[-1])
                except (IndexError, ValueError):
                    level = 1
                lines.append("#" * min(level, 6) + " " + text)
            elif style == "list paragraph":
                lines.append("- " + text)
            else:
                md_line = _para_to_md(para)
                lines.append(md_line)

        with open(output_path, "w", encoding="utf-8") as f:
            f.write("\n".join(lines))
        return True, f"✓ 完成：{os.path.basename(output_path)}"
    except Exception as e:
        return False, f"✗ 转换失败：{str(e)}"


def _para_to_md(para):
    """将 python-docx 段落对象转为带行内格式的 Markdown 字符串。"""
    result = []
    for run in para.runs:
        text = run.text
        if run.bold and run.italic:
            text = f"***{text}***"
        elif run.bold:
            text = f"**{text}**"
        elif run.italic:
            text = f"*{text}*"
        result.append(text)
    return "".join(result)
