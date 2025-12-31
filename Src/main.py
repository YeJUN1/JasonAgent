from pathlib import Path
from typing import List, Tuple

from pdf_reader import extract_text_from_pdf

try:
    from docx import Document
except ImportError:
    Document = None

BASE_DIR = Path(__file__).resolve().parent.parent
INPUT_DIR = BASE_DIR / "Input" / "data"
TEXT_DIR = BASE_DIR / "Output" / "text"
OUTPUT_DOCX = BASE_DIR / "Output" / "综合文档.docx"


def list_pdf_files(input_dir: Path) -> List[Path]:
    if not input_dir.exists():
        print(f"❌ 输入目录不存在: {input_dir}")
        return []

    pdf_files = [
        path for path in input_dir.iterdir()
        if path.is_file() and path.suffix.lower() == ".pdf"
    ]
    if not pdf_files:
        print(f"❌ 没有找到 PDF 文件，请放入目录: {input_dir}")
    return sorted(pdf_files, key=lambda p: p.name.lower())


def read_text_from_folder(folder: Path) -> str:
    if not folder.exists():
        return ""

    def page_sort_key(path: Path) -> int:
        name = path.stem
        if name.startswith("page_"):
            suffix = name.split("_", 1)[1]
            if suffix.isdigit():
                return int(suffix)
        return 0

    page_files = sorted(folder.glob("page_*.txt"), key=page_sort_key)
    contents = []
    for page_file in page_files:
        with open(page_file, "r", encoding="utf-8", errors="ignore") as f:
            contents.append(f.read())
    return "\n".join(contents).strip()


def write_word_doc(text_blocks: List[Tuple[str, str]], output_path: Path) -> None:
    if Document is None:
        print("❌ 缺少依赖 python-docx，请先安装: pip install python-docx")
        return

    output_path.parent.mkdir(parents=True, exist_ok=True)
    doc = Document()

    for index, (title, text) in enumerate(text_blocks, start=1):
        doc.add_heading(f"{index}. {title}", level=1)
        if text.strip():
            for line in text.splitlines():
                doc.add_paragraph(line)
        else:
            doc.add_paragraph("[空白]")

    doc.save(output_path)
    print(f"✅ Word 已生成: {output_path}")


def main() -> None:
    pdf_files = list_pdf_files(INPUT_DIR)
    if not pdf_files:
        return

    TEXT_DIR.mkdir(parents=True, exist_ok=True)
    collected_texts: List[Tuple[str, str]] = []

    for pdf_file in pdf_files:
        output_folder = TEXT_DIR / pdf_file.stem
        extract_text_from_pdf(str(pdf_file), str(output_folder))
        extracted_text = read_text_from_folder(output_folder)
        collected_texts.append((pdf_file.name, extracted_text))

    write_word_doc(collected_texts, OUTPUT_DOCX)


if __name__ == "__main__":
    main()
    

