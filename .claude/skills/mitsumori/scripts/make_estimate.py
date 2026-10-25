#!/usr/bin/env python3
"""株式会社ホームクレバー 御見積書PDF生成 (川嶋淳子様邸の書式)

使い方: python3 make_estimate.py 入力.json 出力.pdf
  → 出力.pdf と同じ名前の .xlsx（Excel版）も作る。
入力JSONの書式は assets/sample_kawashima.json を参照。
"""
import json
import os
import sys

from reportlab.lib.pagesizes import A4
from reportlab.pdfbase import pdfmetrics
from reportlab.pdfbase.ttfonts import TTFont
from reportlab.pdfgen import canvas

HERE = os.path.dirname(os.path.abspath(__file__))
COMPANY = json.load(open(os.path.join(HERE, "..", "assets", "company.json"), encoding="utf-8"))

FONT = "IPAG"
pdfmetrics.registerFont(TTFont(FONT, os.path.join(HERE, "..", "assets", "ipag.ttf")))

W, H = A4
L, R = 42, 553          # 表の左右端
GRAY = (0.9, 0.9, 0.9)
LIGHT = (0.95, 0.95, 0.95)
COLS = [L, 258, 300, 342, 412, R]   # 名称 | 数量 | 単位 | 金額 | 備考
ROW = 17.5


BLANK = False   # 雛形モード: 金額欄を空欄で出す


def yen(n):
    return "" if BLANK else f"{n:,}"


def wrap(text, width, size):
    lines, cur = [], ""
    for ch in text:
        if pdfmetrics.stringWidth(cur + ch, FONT, size) > width:
            lines.append(cur)
            cur = ch
        else:
            cur += ch
    if cur:
        lines.append(cur)
    return lines or [""]


def compute(data):
    sections, total = [], 0
    for s in data["sections"]:
        items = []
        for it in s["items"]:
            amt = it.get("amount")
            if amt is None:
                amt = 0 if BLANK else round(it["qty"] * it["unit_price"])
            items.append({**it, "amount": amt})
        sub = sum(i["amount"] for i in items)
        sections.append({"name": s["name"], "items": items, "subtotal": sub})
        total += sub
    tax = total * COMPANY["tax_rate_percent"] // 100          # 端数切り捨て
    gross = total + tax
    unit = data.get("round_unit", COMPANY["round_unit"])      # 税込合計を切り捨てる単位
    discount = gross % unit if unit else 0
    return sections, total, tax, discount, gross - discount


class Page:
    def __init__(self, c):
        self.c = c

    def text(self, x, y, s, size=9, bold=False, align="l"):
        c = self.c
        c.setFont(FONT, size)
        if bold:   # IPAGothicに太字がないので重ね打ちで代用
            for dx in (0, 0.35):
                self._draw(x + dx, y, s, align)
        else:
            self._draw(x, y, s, align)

    def _draw(self, x, y, s, align):
        {"l": self.c.drawString, "r": self.c.drawRightString, "c": self.c.drawCentredString}[align](x, y, s)

    def rect(self, x, y, w, h, fill=None, stroke=True, lw=0.6):
        c = self.c
        c.setLineWidth(lw)
        if fill:
            c.setFillColorRGB(*fill)
        c.rect(x, y, w, h, stroke=1 if stroke else 0, fill=1 if fill else 0)
        c.setFillColorRGB(0, 0, 0)

    def line(self, x1, y1, x2, y2, lw=0.6):
        self.c.setLineWidth(lw)
        self.c.line(x1, y1, x2, y2)


def table_header(p, y):
    p.rect(L, y - ROW, R - L, ROW, fill=GRAY)
    for x in COLS[1:-1]:
        p.line(x, y, x, y - ROW)
    for i, (label, a) in enumerate(zip(["名称・内訳", "数量", "単位", "金額（税抜）", "備考"], "ccccc")):
        p.text((COLS[i] + COLS[i + 1]) / 2, y - 12, label, 8.5, align="c")
    return y - ROW


def draw_row(p, y, name, qty, unit, amount, note, section=False):
    nl = wrap(name, COLS[1] - COLS[0] - 14, 8.5) if not section else [name]
    kl = wrap(note, COLS[5] - COLS[4] - 8, 7.5) if note else [""]
    h = max(ROW, 6 + 10.5 * max(len(nl), len(kl)))
    p.rect(L, y - h, R - L, h, fill=LIGHT if section else None)
    for x in COLS[1:-1]:
        p.line(x, y, x, y - h)
    ty = y - 12
    for k, ln in enumerate(nl):
        p.text(COLS[0] + (6 if section else 14), ty - 10.5 * k, ln, 8.5, bold=section)
    if qty != "":
        p.text((COLS[1] + COLS[2]) / 2, ty, str(qty), 8.5, align="c")
        p.text((COLS[2] + COLS[3]) / 2, ty, unit, 8.5, align="c")
    if amount != "":
        p.text(COLS[4] - 6, ty, yen(amount), 8.5, bold=section, align="r")
    for k, ln in enumerate(kl):
        p.text(COLS[4] + 4, ty - 10.5 * k, ln, 7.5)
    return y - h


def draw_page_head(p, data, first=True):
    p.text(W / 2, H - 60, data.get("title", "御 見 積 書"), 16, bold=True, align="c")
    p.line(L, H - 70, R, H - 70, 1.6)
    if not first:
        return H - 90
    cust = f'{data["customer"]} {data.get("honorific", "様")}'
    p.text(L + 2, H - 100, cust, 12)
    p.line(L, H - 107, 270, H - 107, 0.6)
    p.text(R, H - 94, data["date"], 8.5, align="r")
    p.text(L, H - 124, data.get("intro", COMPANY["default_intro"]), 8.5)
    return H - 135


def build(data, out):
    sections, subtotal, tax, discount, total = compute(data)
    c = canvas.Canvas(out, pagesize=A4)
    c.setTitle(f'御見積書 {data["customer"]}様')
    p = Page(c)
    y = draw_page_head(p, data)

    # 御見積金額
    p.rect(L + 1, y - 38, 322, 33, stroke=True, lw=1.6)
    p.rect(L + 1, y - 38, 62, 33, fill=GRAY, lw=0.6)
    p.text(L + 32, y - 25, "御見積金額", 8.5, bold=True, align="c")
    p.line(L + 66, y - 12, L + 66, y - 32, 1.4)
    p.text(L + 76, y - 28, f"￥{yen(total)}" + ("" if BLANK else "－"), 17, bold=True)
    p.text(L + 260, y - 27, "（消費税込）", 7.5)

    # 工事情報
    y -= 52
    info = [("工事名称", data["project"]), ("工事場所", data["place"]), ("工期", data["period"]),
            ("お支払い条件", data["payment"]), ("有効期限", data.get("valid", COMPANY["default_valid"]))]
    top, rows = y, []
    for k, v in info:
        ls = wrap(v, 178, 8)
        rows.append((k, ls, max(17, 6 + 10 * len(ls))))
    total_h = sum(r[2] for r in rows)
    p.rect(L, top - total_h, 275, total_h)
    yy = top
    for k, ls, h in rows:
        p.rect(L, yy - h, 78, h, fill=GRAY)
        p.line(L, yy - h, L + 275, yy - h)
        p.text(L + 4, yy - 12, k, 8, bold=True)
        for i, ln in enumerate(ls):
            p.text(L + 82, yy - 12 - 10 * i, ln, 8)
        yy -= h

    # 会社情報
    cx = 320
    p.text(cx, top - 8, COMPANY["name"], 10.5, bold=True)
    p.text(cx, top - 24, f'{COMPANY["rep_title"]}　{COMPANY["rep_name"]}', 8.5)
    cy = top - 38
    for ln in COMPANY["address_lines"]:
        p.text(cx, cy, ln, 7); cy -= 10.5
    c.setDash(2, 2)
    p.rect(504, top - 40, 49, 38, lw=0.5)
    c.setDash()
    p.text(528, top - 24, "印", 7, align="c")
    c.setFillColorRGB(0, 0, 0)
    for ln in COMPANY["license_lines"]:
        p.text(cx, cy, ln, 7); cy -= 10.5
    staff = data.get("staff", COMPANY["default_staff"])
    p.text(cx, cy - 2, f"担当：{staff}", 8)

    # 明細表
    y = min(top - total_h, cy) - 16
    y = table_header(p, y)
    bottom = 50

    def ensure(y, need):
        if y - need < bottom:
            c.showPage()
            y = draw_page_head(p, data, first=False)
            y = table_header(p, y)
        return y

    for s in sections:
        y = ensure(y, ROW * 2)
        y = draw_row(p, y, f'{s["name"]}', "", "", s["subtotal"], "", section=True)
        for it in s["items"]:
            y = ensure(y, ROW + 10)
            y = draw_row(p, y, it["name"], it.get("qty", ""), it.get("unit", ""), it["amount"], it.get("note", ""))

    # 合計欄
    y = ensure(y, ROW * 5)
    def total_row(y, label, val, big=False, fill=None):
        h = 21 if big else ROW
        p.rect(L, y - h, R - L, h, fill=fill)
        p.text(COLS[3] - 6, y - (14 if big else 12), label, 10 if big else 8.5, bold=big, align="r")
        p.text(COLS[4] - 6, y - (14 if big else 12), val, 10 if big else 8.5, bold=big, align="r")
        p.line(COLS[3], y, COLS[3], y - h)
        return y - h
    y = total_row(y, "小計", yen(subtotal))
    y = total_row(y, f'消費税（{COMPANY["tax_rate_percent"]}%）', yen(tax))
    if discount or BLANK:
        y = total_row(y, "端数値引き", "" if BLANK else f"－{yen(discount)}")
    y = total_row(y, "合計（税込）", yen(total), big=True, fill=GRAY)

    # 注記
    y -= 14
    for n in data.get("notes", []):
        ls = wrap("※ " + n, R - L - 8, 8)
        for i, ln in enumerate(ls):
            pass
            p.text(L + (0 if i == 0 else 8), y, ln, 8)
            y -= 11
    c.showPage()
    c.save()
    return total


if __name__ == "__main__":
    if len(sys.argv) != 3:
        sys.exit("usage: make_estimate.py input.json output.pdf")
    d = json.load(open(sys.argv[1], encoding="utf-8"))
    BLANK = bool(d.get("blank"))
    t = build(d, sys.argv[2])
    from excel import build_xlsx
    xlsx = os.path.splitext(sys.argv[2])[0] + ".xlsx"
    build_xlsx(d, COMPANY, compute(d)[0], xlsx, blank=BLANK)
    print(f"OK: {sys.argv[2]} / {xlsx}  合計(税込) ¥{t:,}")
