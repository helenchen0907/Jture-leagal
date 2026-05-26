#!/usr/bin/env python3
"""
按收信人名字拆分扫描型 PDF（无文字层）。

策略：
- 对每封信的第一页，只渲染左上角"地址块"区域为图片
- 用 Tesseract OCR 这一小块，提取收信人名字
- 用 pypdf 把原 PDF 的对应页直接复制到输出文件（保留原画质）

依赖：
    pip install pymupdf pytesseract pillow pypdf
    Tesseract 程序：https://github.com/UB-Mannheim/tesseract/wiki
    （安装时记得勾选"Add to PATH"；或在下面 TESSERACT_CMD 改成实际路径）

用法：
    python split_pdf_ocr.py big_letters.pdf -o output
"""
from __future__ import annotations

import argparse
import os
import re
import sys
from pathlib import Path

import fitz  # PyMuPDF
import pytesseract
from PIL import Image
from pypdf import PdfReader, PdfWriter


# 如果 Tesseract 没加到 PATH，把下面这行的注释去掉，改成你机器上的实际路径
# pytesseract.pytesseract.tesseract_cmd = r"C:\Program Files\Tesseract-OCR\tesseract.exe"


# 抬头关键字，匹配则跳过
HEADER_PATTERNS = [
    r"SUNSHINE\s+EMPIRE",
    r"COMPULSORY\s+LIQUIDATION",
    r"Co\.?\s*Reg",
    r"PricewaterhouseCoopers",
    r"Straits\s+View",
    r"East\s+Tower",
    r"Marina\s+One",
    r"Singapore\s+\d",
    r"Telephone",
    r"Facsimile",
    r"PTE\.?\s*LTD",
]
HEADER_RE = re.compile("|".join(HEADER_PATTERNS), re.IGNORECASE)


def ocr_top_left(page, dpi: int = 300) -> str:
    """渲染页面左上 60% × 65% 区域并 OCR，返回纯文本。"""
    rect = page.rect
    clip = fitz.Rect(0, 0, rect.width * 0.6, rect.height * 0.65)
    pix = page.get_pixmap(dpi=dpi, clip=clip)
    img = Image.frombytes("RGB", (pix.width, pix.height), pix.samples)
    # --psm 6: 假设是一段统一的文字块；只识别常见字母数字标点
    text = pytesseract.image_to_string(img, lang="eng", config="--psm 6")
    return text


def extract_recipient_name(ocr_text: str) -> str:
    """从 OCR 文本里找收信人名字。"""
    for raw_line in ocr_text.splitlines():
        line = raw_line.strip()
        if not line:
            continue
        if HEADER_RE.search(line):
            continue
        if re.search(r"\d", line):
            continue
        # 去掉行内多余空格
        line = re.sub(r"\s+", " ", line)
        # 全大写英文名（允许空格、点、连字符、撇号、斜杠、括号、&）
        if re.fullmatch(r"[A-Z][A-Z\s\.\-'/()&]{1,}", line):
            # 长度过滤：至少 2 个字母
            if len(re.findall(r"[A-Z]", line)) >= 2:
                return line
    return ""


def safe_filename(name: str) -> str:
    cleaned = re.sub(r"[^\w\s\-]", "", name).strip()
    cleaned = re.sub(r"\s+", "_", cleaned)
    return cleaned or "UNKNOWN"


def check_tesseract() -> None:
    try:
        version = pytesseract.get_tesseract_version()
        print(f"Tesseract 版本：{version}")
    except Exception as e:
        print("❌ 找不到 Tesseract 程序。请先安装：")
        print("   https://github.com/UB-Mannheim/tesseract/wiki")
        print("   安装时勾选 'Add to PATH'，或在脚本顶部设置 tesseract_cmd")
        print(f"   原始错误：{e}")
        sys.exit(1)


def split_pdf(input_path: Path, output_dir: Path, pages_per_letter: int = 2,
              dpi: int = 300, debug: bool = False) -> None:
    output_dir.mkdir(parents=True, exist_ok=True)
    reader = PdfReader(str(input_path))
    total = len(reader.pages)

    if total % pages_per_letter != 0:
        print(f"⚠ 警告：总页数 {total} 不能被 {pages_per_letter} 整除")

    used_names: dict[str, int] = {}
    success = 0
    failed: list[tuple[int, str]] = []

    doc = fitz.open(str(input_path))
    try:
        for idx, i in enumerate(range(0, total, pages_per_letter), start=1):
            page = doc[i]
            ocr_text = ocr_top_left(page, dpi=dpi)
            name = extract_recipient_name(ocr_text)

            if debug or not name:
                print(f"\n--- 第 {i+1} 页 OCR 结果 ---")
                print(ocr_text.strip()[:400])
                print("---")

            if not name:
                name = f"UNKNOWN_{idx}"
                failed.append((i + 1, ocr_text.strip()[:80]))

            base = safe_filename(name)
            count = used_names.get(base, 0) + 1
            used_names[base] = count
            filename = f"{base}.pdf" if count == 1 else f"{base}_{count}.pdf"
            out_path = output_dir / filename

            writer = PdfWriter()
            for j in range(pages_per_letter):
                if i + j < total:
                    writer.add_page(reader.pages[i + j])
            with open(out_path, "wb") as f:
                writer.write(f)

            print(f"[{idx}/{total // pages_per_letter}] 第 {i+1}-{i+pages_per_letter} 页 -> {filename}")
            success += 1
    finally:
        doc.close()

    print(f"\n✅ 完成：拆分 {success} 封信到 {output_dir}/")
    if failed:
        print(f"⚠ {len(failed)} 封未识别（用 UNKNOWN_* 命名）：")
        for pno, snippet in failed[:20]:
            print(f"  第 {pno} 页：{snippet!r}")


def main() -> None:
    parser = argparse.ArgumentParser(description="OCR + 按收信人名字拆分 PDF")
    parser.add_argument("input", type=Path)
    parser.add_argument("-o", "--output", type=Path, default=Path("output"))
    parser.add_argument("-n", "--pages", type=int, default=2,
                        help="每封信的页数（默认 2）")
    parser.add_argument("--dpi", type=int, default=300,
                        help="OCR 渲染 DPI（越高越准但越慢，默认 300）")
    parser.add_argument("--debug", action="store_true",
                        help="打印每页的 OCR 结果用于调试")
    args = parser.parse_args()

    if not args.input.exists():
        parser.error(f"输入文件不存在：{args.input}")

    check_tesseract()
    split_pdf(args.input, args.output, args.pages, args.dpi, args.debug)


if __name__ == "__main__":
    main()
