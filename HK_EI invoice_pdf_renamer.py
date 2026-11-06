# -*- coding: utf-8 -*-
"""
发票 PDF 扫描 + 重命名工具（Expeditors 等格式）

功能：
1. 弹窗选择目录（只处理该目录下的 PDF，不包含子文件夹）
2. 扫描每个 PDF，提取 INVOICE NUMBER（如 F916680037）
3. 模式选择：
   - 是(Yes)  = 仅扫描预览：不改文件，只生成扫描报告
   - 否(No)   = 扫描并重命名
   - 取消     = 退出
4. 重复发票号全部保留，不删除、不覆盖：
      F916680037.pdf
      F916680037_DUP2.pdf
      F916680037_DUP3.pdf
5. 每次运行在该目录生成 CSV 日志（原文件名 -> 新文件名），可用于核对或手工还原

依赖：
    pip install pymupdf
扫描件(无文本层)可选 OCR：
    pip install pillow pytesseract   并安装 Tesseract-OCR 程序
"""

import csv
import io
import re
from collections import Counter
from datetime import datetime
from pathlib import Path
import tkinter as tk
from tkinter import filedialog, messagebox

try:
    import fitz  # PyMuPDF
except ImportError:
    raise SystemExit("缺少依赖：请先运行  pip install pymupdf")

# ===== 可调整参数 =====
# 发票号格式：F + 9 位数字。如有其他格式可在此调整，例如 r"[A-Z]\d{9}"
INVOICE_NO_REGEX = r"F\d{9}"
DUP_TAG = "_DUP"          # 重复文件后缀：F916680037_DUP2.pdf
# ======================

PAT_LABEL = re.compile(r"INVOICE\s*NUMBER[\s:#\-]*(" + INVOICE_NO_REGEX + r")\b", re.I)
PAT_ANY = re.compile(r"\b(" + INVOICE_NO_REGEX + r")\b", re.I)


def find_invoice_no(text: str):
    """优先找 'INVOICE NUMBER' 后面的号码；找不到再找全文第一个符合格式的号码。"""
    if not text:
        return None
    m = PAT_LABEL.search(text) or PAT_ANY.search(text)
    return m.group(1).upper() if m else None


def read_text(pdf: Path) -> str:
    with fitz.open(pdf) as doc:
        # 发票号在第 1 页；只读第 1 页可避免误抓后面附件(AWB 等)里的号码
        return doc[0].get_text("text") if len(doc) else ""


def ocr_text(pdf: Path) -> str:
    try:
        import pytesseract
        from PIL import Image
        with fitz.open(pdf) as doc:
            pix = doc[0].get_pixmap(matrix=fitz.Matrix(3, 3), alpha=False)
            img = Image.open(io.BytesIO(pix.tobytes("png")))
            return pytesseract.image_to_string(img, config="--psm 6")
    except Exception:
        return ""


def extract(pdf: Path):
    """返回 (发票号, 识别方式)"""
    no = find_invoice_no(read_text(pdf))
    if no:
        return no, "TEXT"
    no = find_invoice_no(ocr_text(pdf))
    if no:
        return no, "OCR"
    return None, ""


def scan_folder(folder: Path):
    pdfs = sorted([p for p in folder.iterdir() if p.is_file() and p.suffix.lower() == ".pdf"],
                  key=lambda p: p.name.lower())
    rows = []
    for p in pdfs:
        try:
            no, method = extract(p)
            rows.append({"old_name": p.name, "invoice_no": no or "", "method": method,
                         "status": "OK" if no else "NOT_FOUND", "new_name": "", "note": ""})
        except Exception as e:
            rows.append({"old_name": p.name, "invoice_no": "", "method": "",
                         "status": "ERROR", "new_name": "", "note": str(e)})
    return rows


def plan_names(folder: Path, rows):
    """分配新文件名；重复发票号依次加 _DUP2/_DUP3，不覆盖任何已存在文件。"""
    counts = Counter(r["invoice_no"] for r in rows if r["invoice_no"])
    old_names = {r["old_name"].lower() for r in rows}
    taken = {p.name.lower() for p in folder.iterdir() if p.is_file()} - old_names
    seen = Counter()

    # 已经是目标名字的文件优先占位，避免被改成 _DUP
    for r in rows:
        if r["invoice_no"] and r["old_name"].lower() == f'{r["invoice_no"]}.pdf'.lower():
            r["new_name"] = r["old_name"]
            taken.add(r["new_name"].lower())
            seen[r["invoice_no"]] += 1

    for r in rows:
        no = r["invoice_no"]
        if not no:
            r["new_name"] = r["old_name"]
            continue
        if r["new_name"]:
            r["note"] = "已是目标名称"
        else:
            while True:
                seen[no] += 1
                name = f"{no}.pdf" if seen[no] == 1 else f"{no}{DUP_TAG}{seen[no]}.pdf"
                if name.lower() not in taken:
                    break
            r["new_name"] = name
            taken.add(name.lower())
        if counts[no] > 1:
            r["note"] = (r["note"] + "; " if r["note"] else "") + f"重复x{counts[no]}，已保留"
    return rows


def do_rename(folder: Path, rows):
    """两步改名（先改临时名再改目标名），避免文件名互换时冲突。"""
    todo = [r for r in rows if r["status"] == "OK" and r["new_name"] != r["old_name"]]
    temps = []
    for i, r in enumerate(todo):
        src = folder / r["old_name"]
        tmp = folder / f"__tmp_rename_{i}_{datetime.now():%H%M%S%f}.pdf"
        try:
            src.rename(tmp)
            temps.append((tmp, r))
        except Exception as e:
            r["status"], r["note"] = "ERROR", f"改名失败: {e}"
            r["new_name"] = r["old_name"]
    for tmp, r in temps:
        try:
            tmp.rename(folder / r["new_name"])
            r["status"] = "RENAMED"
        except Exception as e:
            tmp.rename(folder / r["old_name"])
            r["status"], r["note"] = "ERROR", f"改名失败: {e}"
            r["new_name"] = r["old_name"]


def write_log(folder: Path, rows, mode: str) -> Path:
    log = folder / f"rename_log_{mode}_{datetime.now():%Y%m%d_%H%M%S}.csv"
    with open(log, "w", newline="", encoding="utf-8-sig") as f:  # utf-8-sig: Excel 打开中文不乱码
        w = csv.DictWriter(f, fieldnames=["old_name", "invoice_no", "method", "status", "new_name", "note"])
        w.writeheader()
        w.writerows(rows)
    return log


def main():
    root = tk.Tk()
    root.withdraw()
    d = filedialog.askdirectory(title="选择 PDF 所在目录（不含子文件夹）")
    if not d:
        return
    folder = Path(d)

    choice = messagebox.askyesnocancel(
        "选择模式",
        "是(Y)：仅扫描预览，不修改文件\n否(N)：扫描并重命名\n取消：退出")
    if choice is None:
        return
    preview = choice

    rows = plan_names(folder, scan_folder(folder))
    if not rows:
        messagebox.showinfo("结果", "该目录下没有 PDF 文件。")
        return

    if not preview:
        do_rename(folder, rows)
    log = write_log(folder, rows, "PREVIEW" if preview else "RENAME")

    st = Counter(r["status"] for r in rows)
    dup = sum(1 for r in rows if "重复" in r["note"])
    for r in rows:
        print(f'[{r["status"]:9}] {r["old_name"]} -> {r["new_name"]}  {r["note"]}')
    messagebox.showinfo(
        "完成",
        f"模式：{'仅扫描预览' if preview else '扫描并重命名'}\n"
        f"PDF 总数：{len(rows)}\n"
        f"识别成功：{len(rows) - st['NOT_FOUND'] - st['ERROR']}\n"
        f"未识别：{st['NOT_FOUND']}（保持原名）\n"
        f"错误：{st['ERROR']}\n"
        f"重复发票文件：{dup}（全部保留）\n\n"
        f"日志：{log.name}")


if __name__ == "__main__":
    main()
