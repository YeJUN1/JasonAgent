from pathlib import Path
from typing import List, Tuple

from pdf_reader import extract_text_from_pdf

BASE_DIR = Path(__file__).resolve().parent.parent
INPUT_DIR = BASE_DIR / "Input" / "data"
TEXT_DIR = BASE_DIR / "Output" / "text"


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


def main() -> None:
    pdf_files = list_pdf_files(INPUT_DIR)
    if not pdf_files:
        return

    TEXT_DIR.mkdir(parents=True, exist_ok=True)

    for pdf_file in pdf_files:
        output_folder = TEXT_DIR / pdf_file.stem
        extract_text_from_pdf(str(pdf_file), str(output_folder))


if __name__ == "__main__":
    main()
    
