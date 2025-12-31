import concurrent.futures
import json
import os
import re
import shutil
import subprocess
from pathlib import Path
from typing import Iterable, List, Optional, Set

from pdf_reader import extract_text_from_pdf
from ocr_client import (
    ocr_image_path_to_text,
    resolve_ocr_workers,
    resolve_visual_ocr_config,
)

try:
    from docx import Document as DocxDocument
except ImportError:
    DocxDocument = None

BASE_DIR = Path(__file__).resolve().parent.parent
INPUT_DIR = BASE_DIR / "Input" / "pdfData"
WORD_INPUT_DIR = BASE_DIR / "Input" / "wordData"
PHOTO_INPUT_DIR = BASE_DIR / "Input" / "photoData"
TEXT_DIR = BASE_DIR / "Output" / "text"
PHOTO_TEXT_DIR = TEXT_DIR / "photo_texts"
OUTPUT_DIR = TEXT_DIR / "combined_documents"
OUTPUT_PDF_TXT = OUTPUT_DIR / "综合文档pdf版本.txt"
OUTPUT_WORD_TXT = OUTPUT_DIR / "综合文档word版本.txt"
OUTPUT_MERGED_TXT = OUTPUT_DIR / "综合合并文档.txt"
OUTPUT_PHOTO_TXT = OUTPUT_DIR / "综合文档图片版本.txt"
SNAPSHOT_FILE = OUTPUT_DIR / "input_snapshot.json"
TITLE_WIDTH = 80
CONTROL_CHARS_RE = re.compile(r"[\x00-\x08\x0b\x0c\x0e-\x1f\x7f]")
HYPERLINK_LINE_RE = re.compile(r"^HYPERLINK\\b", re.IGNORECASE)
PAGE_FIELD_RE = re.compile(r"^PAGE/NUMPAGES$", re.IGNORECASE)
IMAGE_EXTS = {".png", ".jpg", ".jpeg", ".bmp", ".tiff", ".tif", ".webp"}


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


def list_word_files(input_dir: Path) -> List[Path]:
    if not input_dir.exists():
        print(f"❌ 输入目录不存在: {input_dir}")
        return []

    word_files = [
        path for path in input_dir.iterdir()
        if path.is_file() and path.suffix.lower() in {".doc", ".docx"}
    ]
    if not word_files:
        print(f"❌ 没有找到 Word 文件（.doc/.docx），请放入目录: {input_dir}")
    return sorted(word_files, key=lambda p: p.name.lower())


def list_photo_files(input_dir: Path) -> List[Path]:
    if not input_dir.exists():
        print(f"❌ 输入目录不存在: {input_dir}")
        return []

    photo_files = [
        path for path in input_dir.iterdir()
        if path.is_file() and path.suffix.lower() in IMAGE_EXTS
    ]
    if not photo_files:
        print(f"❌ 没有找到图片文件，请放入目录: {input_dir}")
    return sorted(photo_files, key=lambda p: p.name.lower())


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


def iter_word_texts(word_path: Path) -> Iterable[str]:
    def clean_text(value: str) -> str:
        return CONTROL_CHARS_RE.sub("", value).strip()

    suffix = word_path.suffix.lower()
    if suffix == ".docx":
        if DocxDocument is None:
            print("❌ 缺少依赖 python-docx，请先安装: pip install python-docx")
            return

        doc = DocxDocument(word_path)
        for paragraph in doc.paragraphs:
            text = clean_text(paragraph.text)
            if text and not HYPERLINK_LINE_RE.match(text) and not PAGE_FIELD_RE.match(text):
                yield text

        for table in doc.tables:
            for row in table.rows:
                cells = []
                for cell in row.cells:
                    cell_text = " ".join(
                        clean_text(p.text) for p in cell.paragraphs if clean_text(p.text)
                    ).strip()
                    cells.append(cell_text)
                row_text = clean_text("\t".join(cells))
                if row_text and not HYPERLINK_LINE_RE.match(row_text) and not PAGE_FIELD_RE.match(row_text):
                    yield row_text
        return

    if suffix == ".doc":
        textutil_path = shutil.which("textutil")
        if not textutil_path:
            print("❌ 无法读取 .doc 文件，请先安装 LibreOffice 或转换为 .docx")
            return
        result = subprocess.run(
            [textutil_path, "-convert", "txt", "-stdout", str(word_path)],
            capture_output=True,
            text=True,
            encoding="utf-8",
            errors="ignore",
        )
        if result.returncode != 0:
            message = result.stderr.strip() or "转换失败"
            print(f"❌ .doc 解析失败: {word_path.name}（{message}）")
            return
        for line in result.stdout.splitlines():
            line = clean_text(line)
            if line and not HYPERLINK_LINE_RE.match(line) and not PAGE_FIELD_RE.match(line):
                yield line
        return

    print(f"❌ 不支持的 Word 格式: {word_path.name}")


def format_title_line(filename: str) -> str:
    title = filename.strip()
    if len(title) >= TITLE_WIDTH:
        return title
    return title.center(TITLE_WIDTH)


def load_env_file(env_path: Path) -> None:
    if not env_path.exists():
        return
    for raw_line in env_path.read_text(encoding="utf-8").splitlines():
        line = raw_line.strip()
        if not line or line.startswith("#") or "=" not in line:
            continue
        key, value = line.split("=", 1)
        key = key.strip()
        value = value.strip().strip('"').strip("'")
        if key and key not in os.environ:
            os.environ[key] = value


def unique_name(base_name: str, used_names: Set[str], fallback: str = "item") -> str:
    base = base_name.strip() or fallback
    candidate = base
    index = 1
    while candidate.lower() in used_names:
        candidate = f"{base}_{index}"
        index += 1
    used_names.add(candidate.lower())
    return candidate


def resolve_output_folder(pdf_file: Path, used_names: Set[str]) -> Path:
    candidate = unique_name(pdf_file.stem, used_names, fallback="pdf")
    return TEXT_DIR / candidate


def resolve_output_text_file(base_name: str, used_names: Set[str], folder: Path) -> Path:
    candidate = unique_name(base_name, used_names, fallback="text")
    return folder / f"{candidate}.txt"


def write_combined_pdf_txt(pdf_files: List[Path]) -> None:
    if not pdf_files:
        return

    TEXT_DIR.mkdir(parents=True, exist_ok=True)
    OUTPUT_DIR.mkdir(parents=True, exist_ok=True)
    used_folder_names: Set[str] = {OUTPUT_DIR.name.lower()}

    with open(OUTPUT_PDF_TXT, "w", encoding="utf-8") as output_file:
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
    print(f"✅ 综合文档 PDF 版本已生成: {OUTPUT_PDF_TXT}")


def write_combined_word_txt(word_files: List[Path]) -> None:
    if not word_files:
        return

    OUTPUT_DIR.mkdir(parents=True, exist_ok=True)

    with open(OUTPUT_WORD_TXT, "w", encoding="utf-8") as output_file:
        for index, word_file in enumerate(word_files):
            if index > 0:
                output_file.write("\n\n")

            output_file.write(f"{format_title_line(word_file.name)}\n")

            first_line = True
            for line in iter_word_texts(word_file):
                if not first_line:
                    output_file.write("\n")
                output_file.write(line)
                first_line = False
    print(f"✅ 综合文档 Word 版本已生成: {OUTPUT_WORD_TXT}")


def ocr_image_to_text(image_path: Path, config: Optional[dict] = None) -> str:
    return ocr_image_path_to_text(image_path, config)


def write_combined_photo_txt(photo_files: List[Path]) -> None:
    if not photo_files:
        return

    config = resolve_visual_ocr_config()
    if not config:
        return

    PHOTO_TEXT_DIR.mkdir(parents=True, exist_ok=True)
    OUTPUT_DIR.mkdir(parents=True, exist_ok=True)
    used_names: Set[str] = set()
    workers = resolve_ocr_workers()

    results: List[str] = [""] * len(photo_files)
    with concurrent.futures.ThreadPoolExecutor(max_workers=workers) as executor:
        future_map = {
            executor.submit(ocr_image_to_text, photo_file, config): index
            for index, photo_file in enumerate(photo_files)
        }
        for future in concurrent.futures.as_completed(future_map):
            index = future_map[future]
            try:
                results[index] = future.result() or ""
            except Exception as exc:
                print(f"❌ OCR 识别失败: {photo_files[index].name}（{exc}）")
                results[index] = ""

    with open(OUTPUT_PHOTO_TXT, "w", encoding="utf-8") as output_file:
        for index, photo_file in enumerate(photo_files):
            if index > 0:
                output_file.write("\n\n")

            output_file.write(f"{format_title_line(photo_file.name)}\n")
            text = results[index]
            if text:
                output_file.write(text)

            output_text_path = resolve_output_text_file(photo_file.stem, used_names, PHOTO_TEXT_DIR)
            output_text_path.write_text(text, encoding="utf-8")

    print(f"✅ 综合文档 图片版本已生成: {OUTPUT_PHOTO_TXT}")


def write_merged_txt(files: List[Path], output_path: Path) -> None:
    OUTPUT_DIR.mkdir(parents=True, exist_ok=True)
    wrote_any = False
    with open(output_path, "w", encoding="utf-8") as output_file:
        for source_path in files:
            if not source_path.exists() or source_path.stat().st_size == 0:
                continue
            if wrote_any:
                output_file.write("\n\n")
            with open(source_path, "r", encoding="utf-8", errors="ignore") as source_file:
                shutil.copyfileobj(source_file, output_file)
            wrote_any = True
    if wrote_any:
        print(f"✅ 综合合并文档已生成: {output_path}")


def load_snapshot(snapshot_path: Path) -> dict:
    if not snapshot_path.exists():
        return {}
    try:
        return json.loads(snapshot_path.read_text(encoding="utf-8"))
    except (json.JSONDecodeError, OSError):
        return {}


def save_snapshot(snapshot: dict, snapshot_path: Path) -> None:
    OUTPUT_DIR.mkdir(parents=True, exist_ok=True)
    snapshot_path.write_text(
        json.dumps(snapshot, ensure_ascii=True, indent=2),
        encoding="utf-8",
    )


def cleanup_outputs() -> None:
    if OUTPUT_DIR.exists():
        shutil.rmtree(OUTPUT_DIR)
    if not TEXT_DIR.exists():
        return
    for entry in TEXT_DIR.iterdir():
        if entry == OUTPUT_DIR:
            continue
        if entry.is_dir():
            shutil.rmtree(entry)
        elif entry.is_file() and entry.suffix.lower() == ".txt":
            entry.unlink()


def main() -> None:
    load_env_file(BASE_DIR / ".env")
    pdf_files = list_pdf_files(INPUT_DIR)
    word_files = list_word_files(WORD_INPUT_DIR)
    photo_files = list_photo_files(PHOTO_INPUT_DIR)
    current_snapshot = {
        "pdf": sorted(file.name for file in pdf_files),
        "word": sorted(file.name for file in word_files),
        "photo": sorted(file.name for file in photo_files),
    }
    previous_snapshot = load_snapshot(SNAPSHOT_FILE)
    if current_snapshot == previous_snapshot:
        print("ℹ️ 输入文件未变更，跳过重新生成。")
        return

    cleanup_outputs()

    if pdf_files:
        write_combined_pdf_txt(pdf_files)
    if word_files:
        write_combined_word_txt(word_files)
    if photo_files:
        write_combined_photo_txt(photo_files)
    write_merged_txt([OUTPUT_PDF_TXT, OUTPUT_WORD_TXT, OUTPUT_PHOTO_TXT], OUTPUT_MERGED_TXT)
    save_snapshot(current_snapshot, SNAPSHOT_FILE)


if __name__ == "__main__":
    main()
    
