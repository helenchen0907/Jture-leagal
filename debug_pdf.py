#!/usr/bin/env python3
"""诊断 PDF：看看第一页能提取到什么文字。"""
import sys
from pathlib import Path

import pdfplumber


def main():
    if len(sys.argv) < 2:
        print("用法：python debug_pdf.py big_letters.pdf [page_number]")
        sys.exit(1)

    pdf_path = Path(sys.argv[1])
    page_num = int(sys.argv[2]) if len(sys.argv) > 2 else 1

    with pdfplumber.open(str(pdf_path)) as pdf:
        print(f"总页数：{len(pdf.pages)}")
        page = pdf.pages[page_num - 1]
        print(f"页面 {page_num} 尺寸：{page.width} x {page.height}\n")

        # 1) 整页纯文本
        text = page.extract_text() or ""
        print("=== extract_text() 输出 ===")
        if text.strip():
            print(text[:1500])
        else:
            print("(空 —— 大概率是扫描件，没有文字层，需要先 OCR)")
        print()

        # 2) 带坐标的单词（前 30 个）
        words = page.extract_words(use_text_flow=True, keep_blank_chars=False)
        print(f"=== extract_words() 共 {len(words)} 个单词，前 30 个 ===")
        for w in words[:30]:
            print(f"  x0={w['x0']:6.1f}  top={w['top']:6.1f}  text={w['text']!r}")

        # 3) 左半页前几行（这是脚本判断收信人的依据）
        print("\n=== 左半页文字（按 y 排序，前 15 行）===")
        left = [w for w in words if w["x0"] < page.width * 0.5]
        if not left:
            print("(左半页没有文字)")
        else:
            lines = []
            for w in sorted(left, key=lambda x: (x["top"], x["x0"])):
                placed = False
                for line in lines:
                    if abs(w["top"] - line[0]["top"]) <= 3:
                        line.append(w)
                        placed = True
                        break
                if not placed:
                    lines.append([w])
            for line in lines[:15]:
                line_text = " ".join(w["text"] for w in sorted(line, key=lambda w: w["x0"]))
                print(f"  top={line[0]['top']:6.1f}  {line_text!r}")


if __name__ == "__main__":
    main()
