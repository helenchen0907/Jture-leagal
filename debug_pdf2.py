#!/usr/bin/env python3
"""用多个 PDF 库试一试，看哪个能提取到文字。"""
import sys
from pathlib import Path


def try_pymupdf(pdf_path: Path, page_num: int):
    print("\n===== 1) PyMuPDF (fitz) =====")
    try:
        import fitz  # PyMuPDF
    except ImportError:
        print("未安装。运行：pip install pymupdf")
        return
    doc = fitz.open(str(pdf_path))
    page = doc[page_num - 1]
    text = page.get_text("text")
    print(f"文字长度：{len(text)}")
    if text.strip():
        print("--- 前 1000 字 ---")
        print(text[:1000])
    else:
        print("(空)")
    # 带坐标的单词
    words = page.get_text("words")  # [x0, y0, x1, y1, text, ...]
    print(f"单词数：{len(words)}")
    for w in words[:20]:
        print(f"  x0={w[0]:6.1f}  y0={w[1]:6.1f}  text={w[4]!r}")
    doc.close()


def try_pdfminer(pdf_path: Path, page_num: int):
    print("\n===== 2) pdfminer.six =====")
    try:
        from pdfminer.high_level import extract_text
    except ImportError:
        print("未安装。运行：pip install pdfminer.six")
        return
    text = extract_text(str(pdf_path), page_numbers=[page_num - 1])
    print(f"文字长度：{len(text)}")
    if text.strip():
        print("--- 前 1000 字 ---")
        print(text[:1000])
    else:
        print("(空)")


def try_pypdf(pdf_path: Path, page_num: int):
    print("\n===== 3) pypdf =====")
    try:
        from pypdf import PdfReader
    except ImportError:
        print("未安装。运行：pip install pypdf")
        return
    reader = PdfReader(str(pdf_path))
    page = reader.pages[page_num - 1]
    text = page.extract_text() or ""
    print(f"文字长度：{len(text)}")
    if text.strip():
        print("--- 前 1000 字 ---")
        print(text[:1000])
    else:
        print("(空)")


def main():
    if len(sys.argv) < 2:
        print("用法：python debug_pdf2.py big_letters.pdf [page_number]")
        sys.exit(1)
    pdf_path = Path(sys.argv[1])
    page_num = int(sys.argv[2]) if len(sys.argv) > 2 else 1
    print(f"测试文件：{pdf_path}  第 {page_num} 页")

    try_pymupdf(pdf_path, page_num)
    try_pdfminer(pdf_path, page_num)
    try_pypdf(pdf_path, page_num)


if __name__ == "__main__":
    main()
