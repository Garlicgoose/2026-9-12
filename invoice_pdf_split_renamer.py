# -*- coding: utf-8 -*-
"""
发票 PDF 扫描 + 拆分 + 重命名工具（Expeditors 格式）

适用场景：
  一个 PDF 里可能有 1 张或多张发票，后面附提单(HAWB/AWB)，例如：
      第1页  发票 F916621018 (SL 01, 主运费)
      第2页  发票 F916637081 (SL 02, DELIVERY)
      第3页  提单 HAWB 4915200927
  处理后：
      F916621018.pdf  = 第1页 + 第3页(提单)
      F916637081.pdf  = 第2页 + 第3页(提单)
  每张发票都会带上该文件里的提单页。

功能：
1. 弹窗选择目录（只处理该目录下的 PDF，不含子文件夹）
2. 模式：是=仅扫描预览(不改任何文件)；否=拆分并重命名；取消=退出
3. 原文件不删除：处理成功后移到该目录下的 "_原文件" 子文件夹
   （单发票文件同样按发票号生成新文件，规则统一）
4. 发票号重复全部保留，不覆盖：F916680037.pdf, F916680037_DUP2.pdf ...
5. 每次运行在该目录生成 CSV 日志，含 原文件 / 发票号 / 发票页 / 提单页 / 新文件名 / HAWB

依赖：pip install pymupdf
扫描件(无文本层)可选 OCR：pip install pillow pytesseract，并安装 Tesseract-OCR
"""

import csv
import io
import re
import shutil
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
INVOICE_NO_REGEX = r"F\d{9}"        # 发票号格式：F + 9 位数字
DUP_TAG = "_DUP"                     # 重复后缀
ORIGINAL_DIR = "_原文件"             # 原文件归档子文件夹名
# 判定为提单/附件页的关键字（满足任一即视为附件页，会附到每张发票后面）
ATTACH_KEYWORDS = ["AIR WAYBILL", "HAWB", "BILL OF LADING", "SHIPPER'S NAME AND ADDRESS"]
# ======================

PAT_INV = re.compile(r"\b(" + INVOICE_NO_REGEX + r")\b", re.I)
PAT_HAWB = re.compile(r"HAWB\s*(\d{10})", re.I)


# ---------- 文本读取 ----------
def ocr_page(page) -> str:
    try:
        import pytesseract
        from PIL import Image
        pix = page.get_pixmap(matrix=fitz.Matrix(3, 3), alpha=False)
        return pytesseract.image_to_string(Image.open(io.BytesIO(pix.tobytes("png"))), config="--psm 6")
    except Exception:
        return ""


def page_text(page) -> str:
    t = page.get_text("text") or ""
    if len(t.strip()) < 20:          # 基本无文字层 -> 尝试 OCR
        t = ocr_page(page) or t
    return t


# ---------- 页面分类 ----------
def classify_pages(doc):
    """
    返回 invoices: [{"no": 发票号, "pages": [页索引...]}], attach_pages: [页索引...], hawb
    规则：
      - 含 'INVOICE NUMBER' 且有发票号、且不是提单页 -> 新发票的起始页
      - 含提单关键字 -> 附件页
      - 其他页：若前面已有发票且尚未出现附件页 -> 视为该发票续页；否则视为附件页
    """
    invoices, attach, hawb = [], [], ""
    seen_attach = False
    for i, page in enumerate(doc):
        t = page_text(page)
        up = t.upper()
        is_attach = any(k in up for k in ATTACH_KEYWORDS)
        m = PAT_INV.search(t) if "INVOICE NUMBER" in up.replace("\n", " ") or "INVOICE" in up else None
        if is_attach:
            attach.append(i)
            seen_attach = True
            h = PAT_HAWB.search(t)
            if h and not hawb:
                hawb = h.group(1)
        elif m:
            invoices.append({"no": m.group(1).upper(), "pages": [i]})
        elif invoices and not seen_attach:
            invoices[-1]["pages"].append(i)
        else:
            attach.append(i)
    return invoices, attach, hawb


# ---------- 扫描 ----------
def scan_folder(folder: Path):
    pdfs = sorted([p for p in folder.iterdir() if p.is_file() and p.suffix.lower() == ".pdf"],
                  key=lambda p: p.name.lower())
    rows = []
    for p in pdfs:
        try:
            with fitz.open(p) as doc:
                invoices, attach, hawb = classify_pages(doc)
            if not invoices:
                rows.append(dict(src=p.name, invoice_no="", inv_pages="", attach_pages="",
                                 hawb=hawb, new_name="", status="NOT_FOUND", note="未识别到发票号，未处理"))
                continue
            n = len(invoices)
            for inv in invoices:
                rows.append(dict(
                    src=p.name, invoice_no=inv["no"],
                    inv_pages=",".join(str(x + 1) for x in inv["pages"]),
                    attach_pages=",".join(str(x + 1) for x in attach),
                    hawb=hawb, new_name="", status="OK",
                    note=f"该文件含{n}张发票" if n > 1 else "",
                    _pages=inv["pages"] + attach))
        except Exception as e:
            rows.append(dict(src=p.name, invoice_no="", inv_pages="", attach_pages="",
                             hawb="", new_name="", status="ERROR", note=str(e)))
    return rows


def plan_names(folder: Path, rows):
    taken = {p.name.lower() for p in folder.iterdir() if p.is_file()}
    # 原文件处理后会被移走，其名字可被新文件使用
    srcs_ok = {r["src"].lower() for r in rows if r["status"] == "OK"}
    taken -= srcs_ok
    counts = Counter(r["invoice_no"] for r in rows if r["status"] == "OK")
    seen = Counter()
    for r in rows:
        if r["status"] != "OK":
            continue
        no = r["invoice_no"]
        while True:
            seen[no] += 1
            name = f"{no}.pdf" if seen[no] == 1 else f"{no}{DUP_TAG}{seen[no]}.pdf"
            if name.lower() not in taken:
                break
        taken.add(name.lower())
        r["new_name"] = name
        if counts[no] > 1:
            r["note"] = (r["note"] + "; " if r["note"] else "") + f"发票号重复x{counts[no]}，已保留"
    return rows


# ---------- 执行 ----------
def execute(folder: Path, rows):
    arch = folder / ORIGINAL_DIR
    arch.mkdir(exist_ok=True)
    by_src = {}
    for r in rows:
        if r["status"] == "OK":
            by_src.setdefault(r["src"], []).append(r)

    for src, items in by_src.items():
        src_path = folder / src
        # 1) 先把原文件移到归档目录（防止新文件名与原文件名相同时冲突）
        dst_arch = arch / src
        k = 2
        while dst_arch.exists():
            dst_arch = arch / f"{Path(src).stem}_{k}.pdf"
            k += 1
        try:
            shutil.move(str(src_path), str(dst_arch))
        except Exception as e:
            for r in items:
                r["status"], r["note"] = "ERROR", f"原文件移动失败: {e}"
            continue
        # 2) 从归档文件拆分出各发票
        try:
            with fitz.open(dst_arch) as doc:
                for r in items:
                    out = fitz.open()
                    for pg in r["_pages"]:
                        out.insert_pdf(doc, from_page=pg, to_page=pg)
                    out.save(folder / r["new_name"], garbage=3, deflate=True)
                    out.close()
                    r["status"] = "DONE"
        except Exception as e:
            for r in items:
                if r["status"] != "DONE":
                    r["status"], r["note"] = "ERROR", f"拆分失败: {e}（原文件在 {ORIGINAL_DIR}）"


def write_log(folder: Path, rows, mode: str) -> Path:
    log = folder / f"split_log_{mode}_{datetime.now():%Y%m%d_%H%M%S}.csv"
    fields = ["src", "invoice_no", "inv_pages", "attach_pages", "hawb", "new_name", "status", "note"]
    head = ["原文件", "发票号", "发票页", "提单/附件页", "HAWB", "新文件名", "状态", "备注"]
    with open(log, "w", newline="", encoding="utf-8-sig") as f:
        w = csv.writer(f)
        w.writerow(head)
        for r in rows:
            w.writerow([r.get(k, "") for k in fields])
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
        "是(Y)：仅扫描预览，不修改任何文件\n"
        "否(N)：拆分并重命名（原文件移到 _原文件 文件夹，不删除）\n"
        "取消：退出")
    if choice is None:
        return
    preview = choice

    rows = plan_names(folder, scan_folder(folder))
    if not rows:
        messagebox.showinfo("结果", "该目录下没有 PDF 文件。")
        return
    if not preview:
        execute(folder, rows)
    log = write_log(folder, rows, "PREVIEW" if preview else "SPLIT")

    for r in rows:
        print(f'[{r["status"]:9}] {r["src"]} p{r["inv_pages"]}+[{r["attach_pages"]}] -> {r["new_name"]}  {r["note"]}')

    files = len({r["src"] for r in rows})
    inv = sum(1 for r in rows if r["invoice_no"])
    multi = len({r["src"] for r in rows if "张发票" in r["note"]})
    st = Counter(r["status"] for r in rows)
    messagebox.showinfo(
        "完成",
        f"模式：{'仅扫描预览' if preview else '拆分并重命名'}\n"
        f"PDF 文件数：{files}\n"
        f"识别到发票：{inv} 张\n"
        f"含多张发票的文件：{multi} 个\n"
        f"未识别：{st['NOT_FOUND']}（保持原样）\n"
        f"错误：{st['ERROR']}\n\n"
        f"日志：{log.name}")


if __name__ == "__main__":
    main()
