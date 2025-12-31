import re
from pathlib import Path
from typing import List, Tuple

try:
    from docx import Document
    from docx.oxml.ns import qn
except ImportError:
    Document = None
    qn = None

HEADER_RE = re.compile(r"^([一二三四五六七八九十]+、|\d+[\.、]|[（(][一二三四五六七八九十0-9]+[)）])")
LABEL_LINE_RE = re.compile(r"^([^：\s]{1,8}：)(.*)$")
LABEL_INLINE_MAX = 20
PARAGRAPH_ENDINGS = ("。", "！", "？", "；", ":", "：", ".", "!", "?", ";")


def is_standalone_line(line: str) -> bool:
    if line.endswith(("：", ":")):
        return True
    if HEADER_RE.match(line):
        return True
    if line.startswith(("•", "-", "*")):
        return True
    return False


def ends_paragraph(line: str) -> bool:
    return line.endswith(PARAGRAPH_ENDINGS)


def join_lines(prev: str, next_line: str) -> str:
    if not prev:
        return next_line
    prev_tail = prev[-1]
    next_head = next_line[0]
    if prev_tail.isascii() and next_head.isascii() and prev_tail.isalnum() and next_head.isalnum():
        return f"{prev} {next_line}"
    return f"{prev}{next_line}"


def normalize_text_to_paragraphs(text: str) -> List[str]:
    paragraphs: List[str] = []
    current = ""
    for raw_line in text.splitlines():
        line = raw_line.strip()
        if not line:
            if current:
                paragraphs.append(current)
                current = ""
            continue
        label_match = LABEL_LINE_RE.match(line)
        if label_match:
            if current:
                paragraphs.append(current)
                current = ""
            rest = label_match.group(2).strip()
            if not rest:
                paragraphs.append(line)
                continue
            if ends_paragraph(line) or len(rest) <= LABEL_INLINE_MAX:
                paragraphs.append(line)
                continue
            current = line
            if ends_paragraph(current):
                paragraphs.append(current)
                current = ""
            continue
        if is_standalone_line(line):
            if current:
                paragraphs.append(current)
                current = ""
            paragraphs.append(line)
            continue
        if current:
            if ends_paragraph(current):
                paragraphs.append(current)
                current = line
            else:
                current = join_lines(current, line)
        else:
            current = line
        if ends_paragraph(current):
            paragraphs.append(current)
            current = ""
    if current:
        paragraphs.append(current)
    return paragraphs


def write_word_doc(text_blocks: List[Tuple[str, str]], output_path: Path) -> None:
    if Document is None:
        print("❌ 缺少依赖 python-docx，请先安装: pip install python-docx")
        return

    output_path.parent.mkdir(parents=True, exist_ok=True)
    doc = Document()
    normal_style = doc.styles["Normal"]
    normal_style.font.bold = False
    normal_style.font.name = "SimSun"
    if qn is not None:
        normal_style._element.rPr.rFonts.set(qn("w:eastAsia"), "SimSun")

    for index, (title, text) in enumerate(text_blocks, start=1):
        title_paragraph = doc.add_paragraph(style="Normal")
        title_run = title_paragraph.add_run(f"{index}. {title}")
        title_run.bold = False
        if text.strip():
            for paragraph_text in normalize_text_to_paragraphs(text):
                paragraph = doc.add_paragraph(style="Normal")
                run = paragraph.add_run(paragraph_text)
                run.bold = False
        else:
            paragraph = doc.add_paragraph(style="Normal")
            run = paragraph.add_run("[空白]")
            run.bold = False

    doc.save(output_path)
    print(f"✅ Word 已生成: {output_path}")
