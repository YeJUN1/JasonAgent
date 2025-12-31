import pdfplumber
from pdf2image import convert_from_path
import pytesseract
import os
import re
from langdetect import detect, DetectorFactory
from collections import Counter

DetectorFactory.seed = 0  # 保持 langdetect 结果稳定


def detect_language(text, min_chars=100):
    """改进的语言检测（字符占比 + langdetect）"""
    text = text or ""

    # 统计字符占比
    chinese_count = len(re.findall(r"[\u4e00-\u9fff]", text))  # 中文字符
    japanese_count = len(re.findall(r"[\u3040-\u30ff]", text))  # 日文字符
    english_count = len(re.findall(r"[a-zA-Z]", text))  # 英文字母
    total_chars = chinese_count + japanese_count + english_count

    # 计算占比
    if total_chars == 0:
        return "en"

    chi_ratio = chinese_count / total_chars
    jpn_ratio = japanese_count / total_chars
    eng_ratio = english_count / total_chars

    # 判断主要语言
    if eng_ratio > 0.8:
        return "en"
    elif chi_ratio > 0.5:
        return "zh-cn"
    elif jpn_ratio > 0.5:
        return "ja"

    if total_chars < min_chars:
        if chi_ratio >= max(eng_ratio, jpn_ratio):
            return "zh-cn"
        if jpn_ratio >= max(eng_ratio, chi_ratio):
            return "ja"
        return "en"

    # 语言混合时才用 langdetect 进一步检测
    results = []
    for _ in range(5):  # 进行5次检测，提高稳定性
        try:
            results.append(detect(text))
        except:
            continue

    if results:
        return Counter(results).most_common(1)[0][0]  # 返回出现最多次的语言

    return "en"

def extract_text_from_pdf(pdf_path, output_folder):
    """自动选择合适方法提取 PDF 文本，并从第5页后判断主要语言"""
    os.makedirs(output_folder, exist_ok=True)

    print(f"📄 解析 PDF 文件: {pdf_path}")

    is_text_pdf = False
    full_text = ""
    text_for_language_detection = ""  # 用于语言检测的文本

    with pdfplumber.open(pdf_path) as pdf:
        if any(page.extract_text() for page in pdf.pages[:3]):
            is_text_pdf = True

        if is_text_pdf:
            print("📄 该 PDF 具有可选文本，使用 pdfplumber 提取...")
            for i, page in enumerate(pdf.pages):
                text = page.extract_text() or ""
                full_text += text + "\n"

                # 只从第 5 页开始统计语言
                if i >= 4:
                    text_for_language_detection += text + "\n"

                with open(f"{output_folder}/page_{i + 1}.txt", "w", encoding="utf-8") as f:
                    f.write(text)

            lang_sample = text_for_language_detection.strip() or full_text
            detected_lang = detect_language(lang_sample)
            # 存入环境变量（适用于当前运行环境）
            os.environ["DETECTED_LANG"] = detected_lang

            # 存入文件，便于其他 Python 文件访问
            lang_file = os.path.join(output_folder, "lang.txt")
            with open(lang_file, "w", encoding="utf-8") as f:
                f.write(detected_lang)
        else:
            print("🖼️ 该 PDF 似乎是影印版，使用 OCR 识别...")
            images = convert_from_path(pdf_path)
            for i, image in enumerate(images):
                text = pytesseract.image_to_string(image, lang="eng")  # 先默认英文
                full_text += text + "\n"

                # 只从第 5 页开始统计语言
                if i >= 4:
                    text_for_language_detection += text + "\n"

                with open(f"{output_folder}/page_{i + 1}.txt", "w", encoding="utf-8") as f:
                    f.write(text)

            lang_sample = text_for_language_detection.strip() or full_text
            detected_lang = detect_language(lang_sample)

            # 存入环境变量（适用于当前运行环境）
            os.environ["DETECTED_LANG"] = detected_lang

            # 存入文件，便于其他 Python 文件访问
            lang_file = os.path.join(output_folder, "lang.txt")
            with open(lang_file, "w", encoding="utf-8") as f:
                f.write(detected_lang)

        print(f"🌍 主要语言检测结果：{detected_lang}")

    return detected_lang

# import pdfplumber
# import os
# from langdetect import DetectorFactory
# from detect_language import detect_language
#
# DetectorFactory.seed = 0  # 使 langdetect 结果稳定
#
#
# def extract_text_from_pdf(pdf_path, output_folder):
#     """从 PDF 提取文本，并检测主要语言"""
#     os.makedirs(output_folder, exist_ok=True)
#
#     print("📄 该 PDF 具有可选文本，使用 pdfplumber 提取...")
#     full_text = ""
#
#     with pdfplumber.open(pdf_path) as pdf:
#         for i, page in enumerate(pdf.pages):
#             text = page.extract_text() or ""
#             full_text += text + "\n"
#             with open(f"{output_folder}/page_{i + 1}.txt", "w", encoding="utf-8") as f:
#                 f.write(text)
#
#     # 语言检测（从第5页开始，如果页数不足5，则检测所有文本）
#     if len(pdf.pages) >= 5:
#         lang_text = full_text[full_text.find(pdf.pages[4].extract_text()):]
#     else:
#         lang_text = full_text
#
#     detected_lang = detect_language(lang_text)
#
#     # 存入环境变量（适用于当前运行环境）
#     os.environ["DETECTED_LANG"] = detected_lang
#
#     # 存入文件，便于其他 Python 文件访问
#     lang_file = os.path.join(output_folder, "lang.txt")
#     with open(lang_file, "w", encoding="utf-8") as f:
#         f.write(detected_lang)
#
#     print(f"🌍 主要语言检测结果：{detected_lang}")
#     print(f"✅ 提取完成，文本已保存至 {output_folder}")
#
#
