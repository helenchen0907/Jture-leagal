#!/usr/bin/env python3
"""
调试单封信的提取结果。

用法：
    python debug_verify.py output/CHEW_SIEW_NGOH.pdf
"""
import sys
from pathlib import Path

import fitz


def main():
    if len(sys.argv) < 2:
        print("用法：python debug_verify.py path/to/letter.pdf")
        sys.exit(1)

    pdf = Path(sys.argv[1])
    if not pdf.exists():
        print(f"找不到文件：{pdf}")
        sys.exit(1)

    doc = fitz.open(str(pdf))
    full_text = "\n".join(p.get_text("text") for p in doc)

    print("=" * 60)
    print("【完整 PDF 文本】")
    print("=" * 60)
    print(full_text)

    print("\n" + "=" * 60)
    print("【按 span 看文字 + 粗体状态】")
    print("=" * 60)
    for pno, page in enumerate(doc, start=1):
        data = page.get_text("dict")
        for block in data.get("blocks", []):
            if block.get("type") != 0:
                continue
            for line in block.get("lines", []):
                for s in line.get("spans", []):
                    text = s.get("text", "")
                    if not text.strip():
                        continue
                    flags = s.get("flags", 0)
                    font = s.get("font", "")
                    is_bold = bool(flags & 16) or any(
                        t in font.lower() for t in ("bold", "black", "heavy")
                    )
                    mark = "★" if is_bold else " "
                    print(f"  P{pno} {mark} font={font!r:30s} text={text!r}")

    doc.close()


if __name__ == "__main__":
    main()
