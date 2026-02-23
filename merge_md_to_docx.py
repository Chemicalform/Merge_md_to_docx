# -*- coding: utf-8 -*-
"""
遍历脚本所在目录及所有子目录，对每个包含 .md 的文件夹单独生成一个 docx文件，
整合该文件夹下的所有.md 文件，按.md文件名划分章节，正文内容两端对齐，章节间分页，页边距为word默认窄边距格式(1.27cm)。
依赖: pip install python-docx
可选: 安装 pandoc 可保留更多 Markdown 格式 (标题、加粗、列表等)
"""
from pathlib import Path
from docx import Document
from docx.enum.text import WD_ALIGN_PARAGRAPH, WD_BREAK
from docx.shared import Cm
import subprocess
import tempfile
import os
import re

def get_md_files(folder):
    folder = Path(folder)
    files = sorted(folder.glob("*.md"), key=lambda p: p.name.lower())
    return [f for f in files if f.is_file()]

def simple_md_to_paragraphs(doc, text, set_justify=True):
    """将 Markdown 文本按段落加入 docx，简单处理 # 标题与空行。"""
    lines = text.replace("\r\n", "\n").replace("\r", "\n").split("\n")
    i = 0
    while i < len(lines):
        line = lines[i]
        stripped = line.strip()
        if not stripped:
            i += 1
            continue
        if stripped.startswith("# "):
            p = doc.add_heading(stripped[2:].strip(), level=1)
            if set_justify:
                try:
                    p.alignment = WD_ALIGN_PARAGRAPH.JUSTIFY
                except Exception:  # 部分主题下标题样式可能不支持
                    pass
        elif stripped.startswith("## "):
            p = doc.add_heading(stripped[3:].strip(), level=2)
            if set_justify:
                try:
                    p.alignment = WD_ALIGN_PARAGRAPH.JUSTIFY
                except Exception:
                    pass
        elif stripped.startswith("### "):
            p = doc.add_heading(stripped[4:].strip(), level=3)
            if set_justify:
                try:
                    p.alignment = WD_ALIGN_PARAGRAPH.JUSTIFY
                except Exception:
                    pass
        elif stripped in ("---", "***", "___"):
            pass  # 水平线可忽略或留空
        else:
            # 合并连续非空行为一段
            block = [stripped]
            i += 1
            while i < len(lines) and lines[i].strip() and not lines[i].strip().startswith("#"):
                block.append(lines[i].strip())
                i += 1
            para_text = " ".join(block)
            # 简单去掉 Markdown 粗体/斜体标记，先匹配双字符避免把列表 "* item" 破坏
            para_text = re.sub(r"\*\*(.+?)\*\*", r"\1", para_text)
            para_text = re.sub(r"__(.+?)__", r"\1", para_text)
            para_text = re.sub(r"(?<!\*)\*([^*]+)\*(?!\*)", r"\1", para_text)
            para_text = re.sub(r"(?<!_)_([^_]+)_(?!_)", r"\1", para_text)
            p = doc.add_paragraph(para_text)
            if set_justify:
                p.alignment = WD_ALIGN_PARAGRAPH.JUSTIFY
            continue
        i += 1

def copy_paragraphs_with_justify(source_doc, target_doc, skip_first_if_equal=None):
    """将 source_doc 的正文段落复制到 target_doc，并设为两端对齐。
    skip_first_if_equal: 若首段文本与此字符串一致则跳过（避免与章节标题重复）。
    """
    for idx, para in enumerate(source_doc.paragraphs):
        if skip_first_if_equal is not None and idx == 0 and para.text.strip() == skip_first_if_equal:
            continue
        if not para.text.strip():
            target_doc.add_paragraph()
            continue
        new_p = target_doc.add_paragraph()
        new_p.alignment = WD_ALIGN_PARAGRAPH.JUSTIFY
        for run in para.runs:
            r = new_p.add_run(run.text)
            r.bold = run.bold
            r.italic = run.italic
            if run.font.size:
                r.font.size = run.font.size
            if run.font.name:
                r.font.name = run.font.name
    for table in source_doc.tables:
        new_table = target_doc.add_table(rows=len(table.rows), cols=len(table.columns))
        for i, row in enumerate(table.rows):
            for j, cell in enumerate(row.cells):
                new_table.rows[i].cells[j].text = cell.text
                for para in new_table.rows[i].cells[j].paragraphs:
                    para.alignment = WD_ALIGN_PARAGRAPH.JUSTIFY
        # 注意：表格仅复制纯文本，不保留单元格内加粗等格式

def merge_with_pandoc(md_path, out_docx_path):
    """用 pandoc 将单个 md 转为临时 docx。"""
    try:
        subprocess.run(
            ["pandoc", str(md_path), "-o", str(out_docx_path), "-f", "markdown", "-t", "docx"],
            check=True, capture_output=True, creationflags=getattr(subprocess, "CREATE_NO_WINDOW", 0)
        )
        return True
    except (FileNotFoundError, subprocess.CalledProcessError):
        return False

def build_docx_for_folder(folder, md_files, has_pandoc):
    """针对一个文件夹内的 md 列表，生成一个 Document 并返回。页边距 1.27cm。"""
    doc = Document()
    margin_cm = Cm(1.27)
    for section in doc.sections:
        section.top_margin = section.bottom_margin = margin_cm
        section.left_margin = section.right_margin = margin_cm
    for idx, md_path in enumerate(md_files):
        if idx > 0:
            p = doc.add_paragraph()
            p.add_run().add_break(WD_BREAK.PAGE)
        chapter_name = md_path.stem
        doc.add_heading(chapter_name, level=1)
        last_para = doc.paragraphs[-1]
        try:
            last_para.alignment = WD_ALIGN_PARAGRAPH.JUSTIFY
        except Exception:
            pass
        if has_pandoc:
            with tempfile.NamedTemporaryFile(suffix=".docx", delete=False) as tmp:
                tmp_path = tmp.name
            try:
                if merge_with_pandoc(md_path, tmp_path):
                    sub_doc = Document(tmp_path)
                    copy_paragraphs_with_justify(sub_doc, doc, skip_first_if_equal=chapter_name)
                else:
                    text = md_path.read_text(encoding="utf-8", errors="replace")
                    simple_md_to_paragraphs(doc, text, set_justify=True)
            finally:
                try:
                    os.unlink(tmp_path)
                except OSError:
                    pass
        else:
            text = md_path.read_text(encoding="utf-8", errors="replace")
            simple_md_to_paragraphs(doc, text, set_justify=True)
        doc.add_paragraph()
    for section in doc.sections:
        section.top_margin = section.bottom_margin = margin_cm
        section.left_margin = section.right_margin = margin_cm
    return doc


def main():
    root = Path(__file__).resolve().parent
    # 收集所有包含 .md 的目录（含 root 及其子目录）
    dirs_with_md = set()
    for f in root.rglob("*.md"):
        if f.is_file():
            dirs_with_md.add(f.parent)
    dirs_with_md = sorted(dirs_with_md, key=lambda p: (len(p.parts), str(p)))

    if not dirs_with_md:
        print("未在脚本目录及子目录下发现任何 .md 文件。")
        return

    has_pandoc = False
    try:
        r = subprocess.run(
            ["pandoc", "--version"],
            capture_output=True,
            creationflags=getattr(subprocess, "CREATE_NO_WINDOW", 0),
        )
        has_pandoc = r.returncode == 0
    except FileNotFoundError:
        pass

    for folder in dirs_with_md:
        md_files = get_md_files(folder)
        if not md_files:
            continue
        doc = build_docx_for_folder(folder, md_files, has_pandoc)
        out_path = folder / f"{folder.name}.docx"
        doc.save(out_path)
        rel = out_path.relative_to(root)
        print(f"已生成: {rel}（{len(md_files)} 个章节）")

    print(f"共处理 {len(dirs_with_md)} 个文件夹。")

if __name__ == "__main__":
    main()
