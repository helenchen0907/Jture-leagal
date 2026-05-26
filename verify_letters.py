#!/usr/bin/env python3
"""
对比每封 PDF 信件的内容与 Excel 表里的数据。

匹配方式：PDF 文件名 = Excel "Name in POD" 列（空格转下划线后）
输出：把 Excel 复制一份，不匹配的单元格高亮成红色，并加一列 Status

用法：
    python verify_letters.py --excel expected_data.xlsx --pdf-dir output/ --out report.xlsx
    python verify_letters.py --excel expected_data.xlsx --sheet expected_data --pdf-dir output/
"""
from __future__ import annotations

import argparse
import re
import shutil
from datetime import datetime
from pathlib import Path

import fitz  # PyMuPDF
from openpyxl import load_workbook
from openpyxl.styles import PatternFill


# Excel 列名 -> 内部 key（如果列名不一样，改这里）
EXCEL_COLS = {
    "name": "Name in POD",
    "addr1": "Vlookup POD address 1",
    "addr2": "Vlookup POD address 2",
    "addr3": "Vlookup POD address 3",
    "country": "Country",
    "pod_date": "Revised POD Date w/o formula",
    "pod_amount": "POD amount (2dp)",
    "first_div": "First Dividend amount (2 dp)",
    "admitted": "Admitted amount (2dp)",
    "final_div": "Final Dividend amount (2dp)",
    "bank_name": "Bank Name",
    "bank_account": "Bank Account Number",
}

FIELD_KIND = {
    "name": "text",
    "addr1": "text",
    "addr2": "text",
    "addr3": "digits",   # 邮编
    "country": "text",
    "pod_date": "date",
    "pod_amount": "amount",
    "first_div": "amount",
    "admitted": "amount",
    "final_div": "amount",
    "bank_name": "text",
    "bank_account": "digits",
}

RED_FILL = PatternFill(start_color="FFFFC7CE", end_color="FFFFC7CE", fill_type="solid")
YELLOW_FILL = PatternFill(start_color="FFFFEB9C", end_color="FFFFEB9C", fill_type="solid")
ORANGE_FILL = PatternFill(start_color="FFFFD580", end_color="FFFFD580", fill_type="solid")


# 期望的粗体状态：True = SGD 应粗体；False = 不应粗体；不在表里则不检查粗体
BOLD_EXPECTED = {
    "pod_amount": False,
    "first_div": False,
    "admitted": True,
    "final_div": True,
}


# ---------- 归一化 ----------

def norm_text(v) -> str:
    if v is None:
        return ""
    return re.sub(r"\s+", " ", str(v)).strip().upper()


def norm_digits(v) -> str:
    if v is None:
        return ""
    return re.sub(r"\D", "", str(v))


def norm_amount(v) -> str:
    if v is None or v == "":
        return ""
    s = str(v).upper().replace("SGD", "").replace(",", "").strip()
    m = re.search(r"-?\d+(?:\.\d+)?", s)
    return f"{float(m.group()):.2f}" if m else ""


MONTHS = {
    "JAN": 1, "JANUARY": 1, "FEB": 2, "FEBRUARY": 2, "MAR": 3, "MARCH": 3,
    "APR": 4, "APRIL": 4, "MAY": 5, "JUN": 6, "JUNE": 6,
    "JUL": 7, "JULY": 7, "AUG": 8, "AUGUST": 8, "SEP": 9, "SEPT": 9, "SEPTEMBER": 9,
    "OCT": 10, "OCTOBER": 10, "NOV": 11, "NOVEMBER": 11, "DEC": 12, "DECEMBER": 12,
}


def norm_date(v) -> str:
    """归一化成 'YYYY-MM-DD' 字符串。"""
    if v is None or v == "":
        return ""
    if isinstance(v, datetime):
        return v.strftime("%Y-%m-%d")
    s = str(v).strip().upper()
    # 例如 "22 FEBRUARY 2022" 或 "22-FEB-2022" 或 "22/02/2022"
    m = re.match(r"(\d{1,2})[\s\-/]+([A-Z]+|\d{1,2})[\s\-/]+(\d{4})", s)
    if m:
        d, mo, y = m.group(1), m.group(2), m.group(3)
        if mo.isdigit():
            month = int(mo)
        else:
            month = MONTHS.get(mo, 0)
        if month:
            return f"{int(y):04d}-{month:02d}-{int(d):02d}"
    # ISO 格式 YYYY-MM-DD
    m = re.match(r"(\d{4})-(\d{2})-(\d{2})", s)
    if m:
        return f"{m.group(1)}-{m.group(2)}-{m.group(3)}"
    return s


def normalize(v, kind: str) -> str:
    if kind == "amount":
        return norm_amount(v)
    if kind == "date":
        return norm_date(v)
    if kind == "digits":
        return norm_digits(v)
    return norm_text(v)


# ---------- PDF 解析 ----------

def safe_filename(name) -> str:
    cleaned = re.sub(r"[^\w\s\-]", "", str(name)).strip()
    cleaned = re.sub(r"\s+", "_", cleaned)
    return cleaned or "UNKNOWN"


def read_pdf_text(pdf_path: Path) -> str:
    doc = fitz.open(str(pdf_path))
    text = "\n".join(page.get_text("text") for page in doc)
    doc.close()
    return text


def _is_span_bold(span: dict) -> bool:
    """判断 PyMuPDF 的 span 是否粗体。"""
    flags = span.get("flags", 0)
    if flags & 16:  # PyMuPDF flag bit 4 = bold
        return True
    font = span.get("font", "").lower()
    return any(tag in font for tag in ("bold", "black", "heavy"))


def get_amount_bold_map(pdf_path: Path) -> dict[str, bool]:
    """扫 PDF 里所有 'SGD XXX' 形式的金额，返回 {归一化金额: 是否粗体}。

    同一金额出现多次时，只要任意一次是粗体就记为粗体（按需可改）。
    """
    result: dict[str, bool] = {}
    doc = fitz.open(str(pdf_path))
    try:
        for page in doc:
            data = page.get_text("dict")
            for block in data.get("blocks", []):
                if block.get("type") != 0:
                    continue
                for line in block.get("lines", []):
                    spans = line.get("spans", [])
                    if not spans:
                        continue
                    line_text = "".join(s.get("text", "") for s in spans)
                    for m in re.finditer(r"SGD\s*([\d,]+\.\d{2})", line_text):
                        amount = m.group(1)
                        start, end = m.span(0)  # 包含 "SGD" 整段
                        # 找跟 [start, end) 区间有交集的 span
                        is_bold = False
                        cursor = 0
                        for s in spans:
                            s_text = s.get("text", "")
                            s_start = cursor
                            s_end = cursor + len(s_text)
                            cursor = s_end
                            if s_end > start and s_start < end:
                                if _is_span_bold(s):
                                    is_bold = True
                                    break
                        norm = norm_amount(amount)
                        # 多次出现：取 OR（任意一次粗体就算粗体）
                        result[norm] = result.get(norm, False) or is_bold
    finally:
        doc.close()
    return result


AMOUNT_RE = r"(?:SGD|S\$)?\s*([\d,]+\.\d{2})"


def extract_fields_from_pdf(text: str, excel_name: str) -> dict:
    """从 PDF 文本里抽各字段。如果找不到，对应 key 缺失。"""
    out: dict[str, str] = {}

    # POD 日期 + POD 金额（同一句话）
    m = re.search(
        r'POD["\'"]?\)\s*dated\s+(\d{1,2}\s+\w+\s+\d{4})\s+in\s+the\s+amount\s+of\s+' + AMOUNT_RE,
        text, re.IGNORECASE,
    )
    if m:
        out["pod_date"] = m.group(1)
        out["pod_amount"] = m.group(2)

    # First Dividend：you have been paid SGD xxx
    m = re.search(r"you\s+have\s+been\s+paid\s+" + AMOUNT_RE, text, re.IGNORECASE)
    if m:
        out["first_div"] = m.group(1)

    # Admitted：Amount of your POD admitted: SGD xxx
    m = re.search(r"POD\s+admitted[:\s]+" + AMOUNT_RE, text, re.IGNORECASE)
    if m:
        out["admitted"] = m.group(1)

    # Final Dividend：Final dividend payable to you: SGD xxx
    m = re.search(r"Final\s+dividend\s+payable[^:]*:\s*" + AMOUNT_RE, text, re.IGNORECASE)
    if m:
        out["final_div"] = m.group(1)

    # Bank Name
    m = re.search(r"Bank\s+Name\s*[:\-]\s*([^\n\r]+)", text, re.IGNORECASE)
    if m:
        out["bank_name"] = m.group(1).strip()

    # Bank Account Number
    m = re.search(r"Account\s+Number\s*[:\-]\s*([\d\-\s]+)", text, re.IGNORECASE)
    if m:
        out["bank_account"] = m.group(1).strip()

    # 地址块：定位收信人名字所在行，后面 3 行依次是 addr1 / addr2 / addr3+country
    lines = [ln.strip() for ln in text.splitlines()]
    name_up = norm_text(excel_name)
    name_idx = -1
    for i, ln in enumerate(lines):
        if name_up and norm_text(ln) == name_up:
            name_idx = i
            break
    if name_idx >= 0:
        after = [ln for ln in lines[name_idx + 1: name_idx + 6] if ln]
        if len(after) >= 1:
            out["addr1"] = after[0]
        if len(after) >= 2:
            out["addr2"] = after[1]
        if len(after) >= 3:
            # 第三行通常是 "523859 SINGAPORE"
            third = after[2]
            m = re.match(r"(\d{4,6})\s+(.+)", third)
            if m:
                out["addr3"] = m.group(1)
                out["country"] = m.group(2)
            else:
                # 退而求其次：纯邮编 / 纯国家
                if re.fullmatch(r"\d{4,6}", third):
                    out["addr3"] = third
                else:
                    out["country"] = third
        # 有些信件邮编和国家分两行
        if "country" not in out and len(after) >= 4:
            out["country"] = after[3]

    # name 字段也填上，方便对比
    out.setdefault("name", excel_name)
    return out


# ---------- 主流程 ----------

def main() -> None:
    parser = argparse.ArgumentParser(description="核对 PDF 信件与 Excel 数据")
    parser.add_argument("--excel", required=True, type=Path, help="Excel 文件路径")
    parser.add_argument("--sheet", default="expected_data", help="工作表名（默认 expected_data）")
    parser.add_argument("--pdf-dir", required=True, type=Path, help="PDF 所在文件夹")
    parser.add_argument("--out", default=Path("verification_report.xlsx"), type=Path,
                        help="输出 Excel 报告路径")
    args = parser.parse_args()

    if not args.excel.exists():
        parser.error(f"Excel 不存在：{args.excel}")
    if not args.pdf_dir.exists():
        parser.error(f"PDF 文件夹不存在：{args.pdf_dir}")

    # 把 Excel 复制一份做报告
    shutil.copy(args.excel, args.out)
    wb = load_workbook(args.out)
    if args.sheet not in wb.sheetnames:
        print(f"❌ 工作表 {args.sheet!r} 不存在。可用工作表：{wb.sheetnames}")
        return
    ws = wb[args.sheet]

    # 找列号（1-based）。匹配时去掉列头前后空格，避免 "Bank Name " 这种意外。
    header_row = next(ws.iter_rows(min_row=1, max_row=1, values_only=False))
    header_map = {}
    for cell in header_row:
        if cell.value is None:
            continue
        key = str(cell.value).strip()
        header_map[key] = cell.column
    col_idx: dict[str, int] = {}
    missing_cols: list[str] = []
    for key, col_name in EXCEL_COLS.items():
        if col_name in header_map:
            col_idx[key] = header_map[col_name]
        else:
            missing_cols.append(col_name)
    if missing_cols:
        print(f"⚠ Excel 缺少列（这些字段会跳过）：{missing_cols}")
        print(f"   实际列头：{list(header_map.keys())}")

    # 在末尾加一列 Status
    status_col = ws.max_column + 1
    ws.cell(row=1, column=status_col, value="Verify Status")

    total = 0
    pdf_missing = 0
    rows_with_mismatch = 0
    field_mismatches: dict[str, int] = {}
    bold_mismatches: dict[str, int] = {}

    for row in ws.iter_rows(min_row=2):
        total += 1
        name_cell = row[col_idx["name"] - 1] if "name" in col_idx else None
        if not name_cell or not name_cell.value:
            continue

        excel_name = str(name_cell.value).strip()
        pdf_path = args.pdf_dir / f"{safe_filename(excel_name)}.pdf"

        if not pdf_path.exists():
            ws.cell(row=row[0].row, column=status_col, value="PDF MISSING").fill = YELLOW_FILL
            pdf_missing += 1
            continue

        pdf_text = read_pdf_text(pdf_path)
        pdf_fields = extract_fields_from_pdf(pdf_text, excel_name)
        bold_map = get_amount_bold_map(pdf_path)

        row_status = []  # 收集本行的问题
        for key, col in col_idx.items():
            excel_val = row[col - 1].value
            pdf_val = pdf_fields.get(key, "")
            kind = FIELD_KIND[key]
            value_ok = normalize(excel_val, kind) == normalize(pdf_val, kind)

            if not value_ok:
                row[col - 1].fill = RED_FILL
                row[col - 1].comment = _make_comment(f"PDF: {pdf_val}")
                field_mismatches[key] = field_mismatches.get(key, 0) + 1
                row_status.append("MISMATCH")
                continue

            # 值正确，再看 SGD 粗体（仅对金额字段）
            if kind == "amount" and key in BOLD_EXPECTED:
                expected = BOLD_EXPECTED[key]
                actual = bold_map.get(norm_amount(pdf_val))
                if actual is None:
                    continue  # 没在 PDF 找到 SGD 段，跳过粗体检查
                if actual != expected:
                    row[col - 1].fill = ORANGE_FILL
                    row[col - 1].comment = _make_comment(
                        f"值正确但加粗错误：期望 {'粗体' if expected else '非粗体'}, "
                        f"实际 {'粗体' if actual else '非粗体'}"
                    )
                    bold_mismatches[key] = bold_mismatches.get(key, 0) + 1
                    row_status.append("BOLD")

        if "MISMATCH" in row_status:
            ws.cell(row=row[0].row, column=status_col, value="MISMATCH").fill = RED_FILL
            rows_with_mismatch += 1
        elif "BOLD" in row_status:
            ws.cell(row=row[0].row, column=status_col, value="BOLD ISSUE").fill = ORANGE_FILL
            rows_with_mismatch += 1
        else:
            ws.cell(row=row[0].row, column=status_col, value="OK")

    wb.save(args.out)

    print(f"\n========== 核对完成 ==========")
    print(f"总记录：{total}")
    print(f"PDF 缺失：{pdf_missing}")
    print(f"有不匹配字段的记录：{rows_with_mismatch}")
    print(f"完全匹配：{total - pdf_missing - rows_with_mismatch}")
    if field_mismatches:
        print(f"\n各字段值不匹配数：")
        for key, cnt in sorted(field_mismatches.items(), key=lambda x: -x[1]):
            print(f"  {EXCEL_COLS[key]}: {cnt}")
    if bold_mismatches:
        print(f"\n各字段 SGD 粗体不匹配数：")
        for key, cnt in sorted(bold_mismatches.items(), key=lambda x: -x[1]):
            expected = "粗体" if BOLD_EXPECTED[key] else "非粗体"
            print(f"  {EXCEL_COLS[key]} (应为{expected}): {cnt}")
    print(f"\n报告：{args.out}")
    print(f"  红色 = 值不匹配 / 橙色 = 值对但 SGD 粗体不对 / 黄色 = PDF 缺失")


def _make_comment(pdf_val):
    """给单元格加批注，显示 PDF 里读到的值。"""
    from openpyxl.comments import Comment
    return Comment(f"PDF: {pdf_val}", "verify")


if __name__ == "__main__":
    main()
