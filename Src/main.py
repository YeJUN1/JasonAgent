from pathlib import Path
from typing import Iterable, List, Set

from pdf_reader import extract_text_from_pdf

BASE_DIR = Path(__file__).resolve().parent.parent
INPUT_DIR = BASE_DIR / "Input" / "pdfData"
TEXT_DIR = BASE_DIR / "Output" / "text"
OUTPUT_TXT = BASE_DIR / "Output" / "综合文档.txt"
TITLE_WIDTH = 80


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


def iter_page_files(folder: Path) -> List[Path]:
    def page_sort_key(path: Path) -> int:
        name = path.stem
        if name.startswith("page_"):
            suffix = name.split("_", 1)[1]
            if suffix.isdigit():
                return int(suffix)
        return 0

    if not folder.exists():
        return []
    page_files = sorted(folder.glob("page_*.txt"), key=page_sort_key)
    return page_files


def iter_page_texts(folder: Path) -> Iterable[str]:
    for page_file in iter_page_files(folder):
        with open(page_file, "r", encoding="utf-8", errors="ignore") as f:
            text = f.read().strip()
        if text:
            yield text


def read_text_from_folder(folder: Path) -> str:
    return "\n".join(iter_page_texts(folder)).strip()


def format_title_line(filename: str) -> str:
    title = filename.strip()
    if len(title) >= TITLE_WIDTH:
        return title
    return title.center(TITLE_WIDTH)


def resolve_output_folder(pdf_file: Path, used_names: Set[str]) -> Path:
    base_name = pdf_file.stem.strip() or "pdf"
    candidate = base_name
    index = 1
    while candidate.lower() in used_names:
        candidate = f"{base_name}_{index}"
        index += 1
    used_names.add(candidate.lower())
    return TEXT_DIR / candidate


def main() -> None:
    pdf_files = list_pdf_files(INPUT_DIR)
    if not pdf_files:
        return

    TEXT_DIR.mkdir(parents=True, exist_ok=True)
    OUTPUT_TXT.parent.mkdir(parents=True, exist_ok=True)
    used_folder_names: Set[str] = set()

    with open(OUTPUT_TXT, "w", encoding="utf-8") as output_file:
        for index, pdf_file in enumerate(pdf_files):
            output_folder = resolve_output_folder(pdf_file, used_folder_names)
            extract_text_from_pdf(str(pdf_file), str(output_folder))

            if index > 0:
                output_file.write("\n\n")

            output_file.write(f"{format_title_line(pdf_file.name)}\n")

            first_page = True
            for page_text in iter_page_texts(output_folder):
                if not first_page:
                    output_file.write("\n")
                output_file.write(page_text)
                first_page = False
    print(f"✅ 综合文档文本已生成: {OUTPUT_TXT}")


if __name__ == "__main__":
    main()
    
