#!/usr/bin/env python3
"""
按收信人名字拆分 PDF。

假设：
- 每封信占 N 页（默认 2 页）。
- 收信人名字在第一页的左侧，是地址块的第一行（全大写）。

用法：
    python split_pdf_by_recipient.py input.pdf -o output_dir
    python split_pdf_by_recipient.py input.pdf -o output_dir -n 2
"""
from __future__ import annotations

import argparse
import re
from pathlib import Path

import pdfplumber
from pypdf import PdfReader, PdfWriter


# 公司/抬头/页脚等需要跳过的关键字（匹配则不是收信人名字）
HEADER_PATTERNS = [
    r"SUNSHINE\s+EMPIRE",
    r"IN\s+COMPULSORY\s+LIQUIDATION",
    r"Co\.?\s*Reg",
    r"PricewaterhouseCoopers",
    r"Straits\s+View",
    r"East\s+Tower",
    r"Marina\s+One",
    r"Singapore\s+\d",
    r"Telephone",
    r"Facsimile",
    r"^Page\b",
]
HEADER_RE = re.compile("|".join(HEADER_PATTERNS), re.IGNORECASE)


def group_words_into_lines(words, y_tolerance: float = 3.0):
    """把单词按 y 坐标聚合成行，返回 [(top, [word, ...]), ...]，按 y 升序。"""
    lines: list[tuple[float, list]] = []
    for w in sorted(words, key=lambda x: (x["top"], x["x0"])):
        placed = False
        for i, (top, ws) in enumerate(lines):
            if abs(w["top"] - top) <= y_tolerance:
                ws.append(w)
                placed = True
                break
        if not placed:
            lines.append((w["top"], [w]))
    return lines


def extract_recipient_name(page) -> str:
    """从信件首页提取收信人名字。

    策略：取页面左半部分、不含数字、全大写、且不属于公司抬头的第一行文本。
    """
    words = page.extract_words(use_text_flow=True, keep_blank_chars=False)
    if not words:
        return ""

    page_width = page.width
    # 只看左半页（排除右上角的 PwC 地址）
    left_words = [w for w in words if w["x0"] < page_width * 0.5]
    lines = group_words_into_lines(left_words)

    for _, ws in lines:
        ws_sorted = sorted(ws, key=lambda w: w["x0"])
        text = " ".join(w["text"] for w in ws_sorted).strip()
        if not text:
            continue
        if HEADER_RE.search(text):
            continue
        # 名字不含数字
        if re.search(r"\d", text):
            continue
        # 全大写字母组成（允许空格、点、连字符、撇号、斜杠）
        if re.fullmatch(r"[A-Z][A-Z\s\.\-'/()&]{1,}", text):
            return re.sub(r"\s+", " ", text).strip()

    return ""


def safe_filename(name: str) -> str:
    """把名字转成安全的文件名。"""
    cleaned = re.sub(r"[^\w\s\-]", "", name, flags=re.UNICODE).strip()
    cleaned = re.sub(r"\s+", "_", cleaned)
    return cleaned or "UNKNOWN"


def split_pdf(input_path: Path, output_dir: Path, pages_per_letter: int = 2) -> None:
    output_dir.mkdir(parents=True, exist_ok=True)
    reader = PdfReader(str(input_path))
    total = len(reader.pages)

    if total % pages_per_letter != 0:
        print(
            f"⚠ 警告：总页数 {total} 不能被 {pages_per_letter} 整除，"
            f"最后一封信可能不完整。"
        )

    used_names: dict[str, int] = {}
    success = 0
    failed: list[int] = []

    with pdfplumber.open(str(input_path)) as plumber_pdf:
        for idx, i in enumerate(range(0, total, pages_per_letter), start=1):
            first_page = plumber_pdf.pages[i]
            name = extract_recipient_name(first_page)

            if not name:
                name = f"UNKNOWN_{idx}"
                failed.append(i + 1)
                print(f"  [!] 第 {i+1} 页未识别到收信人名字，使用 {name}")

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

            print(f"  第 {i+1}-{i+pages_per_letter} 页 -> {filename}")
            success += 1

    print(f"\n完成：共拆分 {success} 封信到 {output_dir}/")
    if failed:
        print(f"未识别名字的起始页：{failed}（已用 UNKNOWN_* 命名，请手动检查）")


def main() -> None:
    parser = argparse.ArgumentParser(description="按收信人名字拆分 PDF")
    parser.add_argument("input", type=Path, help="输入 PDF 文件")
    parser.add_argument(
        "-o", "--output", type=Path, default=Path("output"),
        help="输出目录（默认: ./output）",
    )
    parser.add_argument(
        "-n", "--pages", type=int, default=2,
        help="每封信的页数（默认: 2）",
    )
    args = parser.parse_args()

    if not args.input.exists():
        parser.error(f"输入文件不存在：{args.input}")

    split_pdf(args.input, args.output, args.pages)


if __name__ == "__main__":
    main()
