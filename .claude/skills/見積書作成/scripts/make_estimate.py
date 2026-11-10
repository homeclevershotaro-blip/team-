#!/usr/bin/env python3
"""御見積書をExcel(+PDF)で作る。使い方:
  python3 -I make_estimate.py 入力.json 出力.xlsx [--pdf]
入力JSONは assets/template_blank.json を参照。company.json は同じ assets フォルダから読む。
target_total を数値で入れると、最初の項目の金額を調整して税込合計をその額に合わせる。"""
import json, sys, subprocess, math
from pathlib import Path
from openpyxl import Workbook
from openpyxl.styles import Font, Alignment, Border, Side, PatternFill
from openpyxl.worksheet.page import PageMargins

FN = "游ゴシック"
thin = Side(style="thin"); med = Side(style="medium")
B = Border(left=thin, right=thin, top=thin, bottom=thin)
G = PatternFill("solid", fgColor="E7E6E6")
L = Alignment(horizontal="left", vertical="center", wrap_text=True)
C = Alignment(horizontal="center", vertical="center", wrap_text=True)
R = Alignment(horizontal="right", vertical="center")
def f(sz=10, b=False): return Font(name=FN, size=sz, bold=b)

def apply_markup(d):
    """markup（例 0.25）があれば、数値の金額に上乗せする（円未満四捨五入）。"""
    m = d.get("markup") or 0
    if not m: return
    for s in d["sections"]:
        for i in s["items"]:
            if isinstance(i["amount"], (int, float)):
                i["amount"] = round(i["amount"] * (1 + m))

def fit_target(d):
    """税込合計を target_total に合わせる（最初の数値項目で調整）。"""
    t = d.get("target_total")
    if not t: return
    net = math.ceil(t / 1.1)
    while net + net // 10 > t: net -= 1
    while net + net // 10 < t: net += 1
    cur = sum(i["amount"] for s in d["sections"] for i in s["items"] if isinstance(i["amount"], (int, float)))
    for s in d["sections"]:
        for i in s["items"]:
            if isinstance(i["amount"], (int, float)):
                i["amount"] += net - cur; return

def build(d, co, out):
    apply_markup(d)
    fit_target(d)
    wb = Workbook(); ws = wb.active; ws.title = "御見積書"
    for c, w in zip("ABCDEF", [40, 7, 7, 15, 40, 2]): ws.column_dimensions[c].width = w
    ws.page_setup.paperSize = 9; ws.page_setup.fitToWidth = 1; ws.page_setup.fitToHeight = 1
    ws.sheet_properties.pageSetUpPr.fitToPage = True
    ws.page_margins = PageMargins(left=.5, right=.5, top=.5, bottom=.5)
    def put(cell, v, font=None, al=None, border=None, fill=None):
        c = ws[cell]; c.value = v; c.font = font or f()
        if al: c.alignment = al
        if border: c.border = border
        if fill: c.fill = fill
    ws.merge_cells("A1:E1"); put("A1", "御 見 積 書", f(20, True), C); ws.row_dimensions[1].height = 34
    put("E2", d["date"], f(), R)
    ws.merge_cells("A3:B3"); put("A3", d["customer"], f(13), L)
    ws.merge_cells("A4:E4"); put("A4", "下記の通り御見積申し上げます。何卒ご用命賜りますようお願い申し上げます。")
    put("A6", "御見積金額", f(11, True), C, B, G)
    ws.merge_cells("B6:D6"); put("B6", None, f(16, True), C); put("E6", "（消費税込）", f(), C)
    for col in "ABCDE":
        ws[col+"6"].border = Border(top=med, bottom=med, left=med if col == "A" else None, right=med if col == "E" else None)
    ws.row_dimensions[6].height = 30
    info = [("工事名称", d["title"]), ("工事場所", d["place"]), ("工期", d["period"]),
            ("お支払い条件", d["payment"]), ("有効期限", d["validity"])]
    for i, (k, v) in enumerate(info):
        r = 8 + i; put(f"A{r}", k, f(10, True), C, B, G); ws.merge_cells(f"B{r}:D{r}")
        put(f"B{r}", v, f(), L, B); ws[f"C{r}"].border = B; ws[f"D{r}"].border = B
        ws.row_dimensions[r].height = 20
    for i, k in enumerate(["name", "rep", "address", "tel", "license", "invoice_no", "staff"]):
        put(f"E{8+i}", co[k], f(12, True) if i == 0 else f(8), Alignment(horizontal="left", vertical="center"))
    for i, h in enumerate(["名称・内訳", "数量", "単位", "金額（税抜）", "備考"]):
        put(f"{'ABCDE'[i]}16", h, f(10, True), C, B, G)
    r = 17; sec_rows = []
    for n, s in enumerate(d["sections"], 1):
        put(f"A{r}", f"{n}. {s['name']}", f(10, True), L, B, G)
        for c in "BCE": put(f"{c}{r}", None, None, None, B, G)
        put(f"D{r}", None, f(10, True), R, B, G); ws[f"D{r}"].number_format = "#,##0"
        first = r + 1; r += 1
        for it in s["items"]:
            put(f"A{r}", "　" + it["name"], f(), L, B); put(f"B{r}", it["qty"], f(), C, B)
            put(f"C{r}", it["unit"], f(), C, B); put(f"D{r}", it["amount"], f(), R, B)
            put(f"E{r}", it.get("note", ""), f(), L, B); ws[f"D{r}"].number_format = "#,##0"
            r += 1
        ws[f"D{first-1}"] = f"=SUM(D{first}:D{r-1})"; sec_rows.append(first - 1)
    r += 1
    def tot(row, label, formula, fill=None):
        ws.merge_cells(f"A{row}:C{row}"); put(f"A{row}", label, f(10, True), R, B, fill)
        for c in "BC": ws[f"{c}{row}"].border = B
        put(f"D{row}", formula, f(10, True), R, B, fill); ws[f"D{row}"].number_format = "#,##0"
        put(f"E{row}", None, None, None, B, fill)
    tot(r, "小計", "=" + "+".join(f"D{x}" for x in sec_rows))
    tot(r+1, "消費税（10%）", f"=ROUNDDOWN(D{r}*0.1,0)")
    tot(r+2, "端数値引き", -abs(d.get("discount", 0)))
    tot(r+3, "合計（税込）", f"=D{r}+D{r+1}+D{r+2}", G)
    ws["B6"] = f'="¥"&TEXT(D{r+3},"#,##0")&"－"'
    for i, t in enumerate(d.get("notes", [])):
        ws.merge_cells(f"A{r+5+i}:E{r+5+i}"); put(f"A{r+5+i}", t, f(9), L)
    ws.print_area = f"A1:E{r+5+len(d.get('notes', []))}"
    wb.save(out)

if __name__ == "__main__":
    a = [x for x in sys.argv[1:] if not x.startswith("--")]
    d = json.load(open(a[0], encoding="utf-8"))
    co = json.load(open(Path(__file__).parent.parent / "assets" / "company.json", encoding="utf-8"))
    build(d, co, a[1])
    if "--pdf" in sys.argv:
        subprocess.run(["soffice", "--headless", "--convert-to", "pdf", "--outdir", str(Path(a[1]).parent), a[1]], check=False)
