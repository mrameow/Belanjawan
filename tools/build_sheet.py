"""Bina Belanjawan_Kelab_Guru_Baru.xlsx (import ke Google Sheets).

Guna:  python3 -I tools/build_sheet.py <tx.json> <output.xlsx>

Tab: Ringkasan | Harian | Transaksi | Yuran Ahli | Tetapan
Transaksi (satu baris = satu bayaran/belanja, seperti helaian lama):
  A ID | B Tarikh | C Ahli/Pihak | D Butiran | E Kategori | F Bulan Yuran |
  G Kelab | H Lawatan | I Masuk Lain | J Keluar | K Baki | L Hari# (pembantu)
"""
import datetime as dt
import json
import sys

from openpyxl import Workbook
from openpyxl.chart import BarChart, LineChart, Reference
from openpyxl.formatting.rule import CellIsRule, FormulaRule
from openpyxl.styles import Alignment, Border, Font, PatternFill, Side
from openpyxl.utils import get_column_letter as CL
from openpyxl.worksheet.datavalidation import DataValidation

TEAL, TEAL_D, TEAL_L = "0F766E", "134E4A", "F0FDFA"
GREEN, GREEN_L, RED, RED_L, AMBER, AMBER_L = "15803D", "DCFCE7", "B91C1C", "FEE2E2", "B45309", "FEF3C7"
GREY, GREY_L, INK = "64748B", "F1F5F9", "0F172A"
FONT = "Arial"
RM = '"RM"#,##0.00;[Red]-"RM"#,##0.00'
RM0 = '#,##0.00;[Red]-#,##0.00;"-"'
MAX = 3000          # baris Transaksi yang disediakan
DAYS = 400          # baris Harian yang disediakan
N_MONTHS = 19       # Jun 2025 - Dis 2026 dalam jadual bulanan / matriks
T = "Transaksi!"

AHLI = ["Amir", "Asnawi", "Bang Mie", "Boss Khairil", "Boss Wan", "Chang", "Hadi",
        "Hasnol", "Joe", "Khairiri", "Kuchai", "Suhaimi", "Ustaz", "Wafi", "Aiman"]
TEMPOH = {  # (Mula, Tamat) anggaran daripada data; boleh diubah di Tetapan
    "Boss Khairil": ((2025, 6), (2025, 9)), "Suhaimi": ((2025, 6), (2025, 8)),
    "Wafi": ((2025, 6), (2026, 3)), "Aiman": ((2025, 12), None),
}
NO_LAWATAN = {"Kuchai": "Tiada kutipan lawatan (refund 400 Jan 2026) - semak",
              "Wafi": "Lawatan diuruskan berasingan - semak"}
KATEGORI = ["Yuran", "Majlis", "Sumbangan", "Lain-lain",                       # masuk
            "Duit Pantri", "Lawatan / Trip", "Majlis & Hadiah", "Pakaian",      # keluar
            "Langganan / Bahan", "Refund", "Pelbagai"]
KAT_KELUAR_ROW = 9  # baris pertama kategori keluar dalam Tetapan (I9..I15)


def fill(c):
    return PatternFill("solid", start_color=c, end_color=c)


def font(**kw):
    kw.setdefault("name", FONT)
    kw.setdefault("size", 10)
    kw.setdefault("color", INK)
    return Font(**kw)


thin = Side(style="thin", color="E2E8F0")
BOX = Border(left=thin, right=thin, top=thin, bottom=thin)
BOT = Border(bottom=thin)


def header(ws, row, labels, col=1, height=28):
    for i, t in enumerate(labels):
        c = ws.cell(row, col + i, t)
        c.font = font(bold=True, color="FFFFFF")
        c.fill = fill(TEAL_D)
        c.alignment = Alignment(horizontal="center", vertical="center", wrap_text=True)
    ws.row_dimensions[row].height = height


def banner(ws, text, sub, last_col):
    ws.sheet_view.showGridLines = False
    for r in (1, 2):
        ws.merge_cells(start_row=r, start_column=1, end_row=r, end_column=last_col)
        for c in range(1, last_col + 1):
            ws.cell(r, c).fill = fill(TEAL)
        ws.cell(r, 1).alignment = Alignment(vertical="center", indent=1)
    ws.cell(1, 1, text).font = font(size=20, bold=True, color="FFFFFF")
    ws.cell(2, 1, sub).font = font(size=10, color="CCFBF1")
    ws.row_dimensions[1].height = 36
    ws.row_dimensions[2].height = 20


def widths(ws, ws_widths, start=1):
    for i, w in enumerate(ws_widths):
        ws.column_dimensions[CL(start + i)].width = w


# ------------------------------------------------------------------ Tetapan
def build_tetapan(ws):
    banner(ws, "Tetapan", "Senarai ahli, kadar yuran dan kategori. Aplikasi & tab lain membaca helaian ini.", 10)
    header(ws, 4, ["Ahli", "Mula", "Tamat", "Kelab / bulan", "Lawatan / bulan", "Aktif", "Catatan"])
    for i, n in enumerate(AHLI):
        r = 5 + i
        mula, tamat = TEMPOH.get(n, ((2025, 6), None))
        vals = [n, dt.date(*mula, 1), dt.date(*tamat, 1) if tamat else None, 35,
                0 if n in NO_LAWATAN else 50, "Tidak" if tamat else "Ya", NO_LAWATAN.get(n)]
        for c, v in enumerate(vals, 1):
            cell = ws.cell(r, c, v)
            cell.font, cell.border = font(), BOX
            if c in (2, 3):
                cell.number_format = "mmm yyyy"
            if c in (4, 5):
                cell.number_format = RM
    ws.cell(21, 1, "Lawatan dikutip sehingga").font = font(bold=True)
    ws.cell(22, 1, "Tarikh mula rekod").font = font(bold=True)
    for r, d in ((21, dt.date(2026, 2, 1)), (22, dt.date(2025, 6, 1))):
        c = ws.cell(r, 2, d)
        c.number_format, c.fill, c.font = "mmm yyyy", fill("FEF9C3"), font()
    ws.cell(24, 1, "Mula/Tamat ahli hanyalah anggaran daripada data lama - semak dan betulkan. "
                   "'Tamat' kosong = masih aktif. Kadar 0 = dikecualikan.").font = font(italic=True, color=GREY)
    header(ws, 4, ["Kategori"], col=9)
    for i, v in enumerate(KATEGORI):
        ws.cell(5 + i, 9, v).font = font()
    ws.cell(5, 10, "← masuk (baris 5-8)").font = font(italic=True, color=GREY, size=9)
    ws.cell(KAT_KELUAR_ROW, 10, "← keluar (baris 9-15)").font = font(italic=True, color=GREY, size=9)
    widths(ws, (22, 12, 12, 15, 16, 8, 52, 3, 22, 22))
    ws.freeze_panes = "A5"


# ---------------------------------------------------------------- Transaksi
def build_transaksi(ws, tx):
    ws.sheet_view.showGridLines = False
    header(ws, 1, ["ID", "Tarikh", "Ahli / Pihak", "Butiran", "Kategori", "Bulan Yuran",
                   "Kelab (RM)", "Lawatan (RM)", "Masuk Lain (RM)", "Keluar (RM)", "Baki (RM)", "Hari #"], height=32)
    for i, x in enumerate(tx):
        vals = [f"M{i + 1:04d}", dt.date.fromisoformat(x["tarikh"]), x["ahli"], x["butiran"], x["kategori"],
                x["bulan"], x["kelab"] or None, x["lawatan"] or None, x["lain"] or None, x["keluar"] or None]
        for c, v in enumerate(vals, 1):
            ws.cell(i + 2, c, v)
    for r in range(2, MAX + 1):
        ws.cell(r, 11, f'=IF(B{r}="","",SUM(G$2:I{r})-SUM(J$2:J{r}))')
        ws.cell(r, 12, f'=IF(B{r}="","",IF(B{r}=B{r - 1},L{r - 1},N(L{r - 1})+1))')
        for c in range(1, 13):
            cell = ws.cell(r, c)
            cell.font = font(color=GREY if c in (1, 6, 11, 12) else INK, size=9 if c in (1, 12) else 10)
            cell.border = BOT
        ws.cell(r, 2).number_format = "ddd, dd mmm yyyy"
        ws.cell(r, 2).alignment = Alignment(horizontal="left")
        for c in range(7, 12):
            ws.cell(r, c).number_format = RM0
    # hari genap berlatar teal muda -> setiap hari nampak terasing
    ws.conditional_formatting.add(f"A2:L{MAX}", FormulaRule(formula=['AND($L2<>"",ISEVEN($L2))'], fill=fill(TEAL_L)))
    # tarikh berulang dalam hari yang sama dipudarkan; baris pertama hari itu tebal
    ws.conditional_formatting.add(f"B2:B{MAX}", FormulaRule(formula=['AND($B2<>"",$B2=$B1)'], font=Font(name=FONT, color="CBD5E1")))
    ws.conditional_formatting.add(f"B2:B{MAX}", FormulaRule(formula=['AND($B2<>"",$B2<>$B1)'], font=Font(name=FONT, bold=True, color=TEAL_D)))
    ws.conditional_formatting.add(f"G2:I{MAX}", CellIsRule(operator="greaterThan", formula=["0"], font=Font(name=FONT, color=GREEN)))
    ws.conditional_formatting.add(f"J2:J{MAX}", CellIsRule(operator="greaterThan", formula=["0"], font=Font(name=FONT, bold=True, color=RED)))
    for col, src in (("E", "$I$5:$I$15"), ("C", "$A$5:$A$19")):
        dv = DataValidation(type="list", formula1=f"=Tetapan!{src}", allow_blank=True, showErrorMessage=False)
        ws.add_data_validation(dv)
        dv.add(f"{col}2:{col}{MAX}")
    widths(ws, (8, 17, 15, 44, 16, 24, 12, 12, 13, 12, 14, 7))
    ws.freeze_panes = "C2"
    ws.auto_filter.ref = f"A1:L{MAX}"
    ws.cell(1, 14, "Bulan Yuran = bulan yang dibayar, format 2026-03, 2026-04. Bayaran separa: 2025-12:30 (RM30).").font = font(italic=True, color=GREY, size=9)


# ------------------------------------------------------------------- Harian
def build_harian(ws):
    banner(ws, "Rekod Harian", "Satu baris = satu hari. Terbaru di atas. Dikemas kini automatik daripada Transaksi.", 8)
    header(ws, 4, ["Hari #", "Tarikh", "Hari", "Bil. rekod", "Masuk", "Keluar", "Bersih", "Baki hujung hari"])
    L = f"{T}$L$2:$L${MAX}"

    def s(cols, op):
        return "+".join(f'SUMIFS({T}${c}$2:${c}${MAX},{L},{op})' for c in cols)
    for k in range(DAYS):
        r = 5 + k
        ws.cell(r, 1, f'=IF({k + 1}<=MAX({L}),MAX({L})-{k},"")')
        ws.cell(r, 2, f'=IF(A{r}="","",INDEX({T}$B$2:$B${MAX},MATCH(A{r},{L},0)))')
        ws.cell(r, 3, f'=IF(A{r}="","",CHOOSE(WEEKDAY(B{r},2),"Isnin","Selasa","Rabu","Khamis","Jumaat","Sabtu","Ahad"))')
        ws.cell(r, 4, f'=IF(A{r}="","",COUNTIF({L},A{r}))')
        ws.cell(r, 5, f'=IF(A{r}="","",{s("GHI", f"A{r}")})')
        ws.cell(r, 6, f'=IF(A{r}="","",{s("J", f"A{r}")})')
        ws.cell(r, 7, f'=IF(A{r}="","",E{r}-F{r})')
        ws.cell(r, 8, f'=IF(A{r}="","",{s("GHI", f"\"<=\"&A{r}")}-({s("J", f"\"<=\"&A{r}")}))')
        for c in range(1, 9):
            cell = ws.cell(r, c)
            cell.font, cell.border = font(color=GREY if c == 1 else INK, size=9 if c == 1 else 10), BOT
            if c >= 5:
                cell.number_format = RM
        ws.cell(r, 2).number_format = "dd mmm yyyy"
        for c in (2, 3):
            ws.cell(r, c).alignment = Alignment(horizontal="left")
        ws.cell(r, 4).alignment = Alignment(horizontal="center")
        ws.cell(r, 5).font, ws.cell(r, 6).font = font(color=GREEN), font(color=RED)
        ws.cell(r, 8).font = font(bold=True)
    last = 4 + DAYS
    ws.conditional_formatting.add(f"A5:H{last}", FormulaRule(formula=['AND($A5<>"",ISEVEN($A5))'], fill=fill(TEAL_L)))
    ws.conditional_formatting.add(f"G5:G{last}", CellIsRule(operator="lessThan", formula=["0"], font=Font(name=FONT, bold=True, color=RED)))
    widths(ws, (7, 14, 10, 10, 15, 15, 15, 18))
    ws.freeze_panes = "A5"


# -------------------------------------------------------------- Yuran Ahli
def matrix(ws, top, title, kind):
    """kind 'K' (kelab, lajur G) atau 'L' (lawatan, lajur H). Mengembalikan baris data pertama."""
    col, rate = ("G", "D") if kind == "K" else ("H", "E")
    ws.cell(top, 1, title).font = font(size=13, bold=True, color=TEAL_D)
    h = top + 1
    header(ws, h, ["Ahli"], height=24)
    for i in range(N_MONTHS):
        c = ws.cell(h, 2 + i, "=Tetapan!$B$22" if i == 0 else f"=EDATE({CL(1 + i)}{h},1)")
        c.number_format, c.font, c.fill = "mmm yy", font(bold=True, color="FFFFFF", size=9), fill(TEAL_D)
        c.alignment = Alignment(horizontal="center", vertical="center")
    first = h + 1
    for k in range(len(AHLI)):
        r, s = first + k, 5 + k
        ws.cell(r, 1, f"=Tetapan!A{s}").font = font(bold=True)
        ws.cell(r, 1).border = BOX
        for i in range(N_MONTHS):
            m = f"{CL(2 + i)}${h}"
            key = f'TEXT({m},"yyyy-mm")'
            who = f"({T}$C$2:$C${MAX}=$A{r})*({T}${col}$2:${col}${MAX}>0)"
            full = f'ISNUMBER(SEARCH(", "&{key}&",",", "&{T}$F$2:$F${MAX}&","))*{who}'
            part = f'ISNUMBER(SEARCH(", "&{key}&":",", "&{T}$F$2:$F${MAX}&","))*{who}'
            out = f'OR({m}<Tetapan!$B{s},AND(Tetapan!$C{s}<>"",{m}>Tetapan!$C{s}),Tetapan!${rate}{s}=0'
            out += f',{m}>Tetapan!$B$21)' if kind == "L" else ")"
            cell = ws.cell(r, 2 + i, f'=IF({out},"–",IF(SUMPRODUCT({full})>0,"✓",IF(SUMPRODUCT({part})>0,"½",IF({m}>TODAY(),"·","✗"))))')
            cell.alignment, cell.border, cell.font = Alignment(horizontal="center"), BOX, font(bold=True, size=11)
        # pembantu (lajur V, tersembunyi): senarai bulan belum bayar
        parts = "&".join(
            f'IF({CL(2 + i)}{r}="✗",TEXT({CL(2 + i)}${h},"mmm yy")&", ",IF({CL(2 + i)}{r}="½",TEXT({CL(2 + i)}${h},"mmm yy")&" (separa), ",""))'
            for i in range(N_MONTHS))
        ws.cell(r, 22, f"={parts}").font = font(color=GREY, size=8)
    rng = f"B{first}:T{first + len(AHLI) - 1}"
    for sym, fc, bg in (("✓", GREEN, GREEN_L), ("✗", RED, RED_L), ("½", AMBER, AMBER_L)):
        ws.conditional_formatting.add(rng, CellIsRule(operator="equal", formula=[f'"{sym}"'], font=Font(name=FONT, bold=True, color=fc), fill=fill(bg)))
    for sym in ("–", "·"):
        ws.conditional_formatting.add(rng, CellIsRule(operator="equal", formula=[f'"{sym}"'], font=Font(name=FONT, color="CBD5E1")))
    return first


def build_yuran(ws):
    banner(ws, "Yuran Ahli", "Bulan mana belum dibayar. Berdasarkan lajur 'Bulan Yuran' di Transaksi dan tempoh di Tetapan.", 20)
    ws.cell(3, 1, "✓ dibayar    ½ bayaran separa    ✗ belum bayar    · bulan akan datang    – di luar tempoh / dikecualikan").font = font(size=9, color=GREY)
    for c in range(1, 21):
        ws.cell(5, c).fill = fill(TEAL_D)
    for c, t in ((1, "Ahli"), (2, "Kelab: bil. belum bayar"), (3, "Kelab: bulan belum bayar"),
                 (11, "Lawatan: bil. belum bayar"), (12, "Lawatan: bulan belum bayar")):
        cell = ws.cell(5, c, t)
        cell.font = font(bold=True, color="FFFFFF")
        cell.alignment = Alignment(horizontal="center", vertical="center", wrap_text=True)
    ws.row_dimensions[5].height = 42
    ws.merge_cells("C5:J5")
    ws.merge_cells("L5:T5")
    n = len(AHLI)
    top_k = 5 + n + 3
    k_first, top_l = top_k + 2, top_k + 2 + n + 2
    l_first = top_l + 2
    for k in range(n):
        r = 6 + k
        kr, lr = k_first + k, l_first + k
        ws.cell(r, 1, f"=Tetapan!A{5 + k}")
        ws.cell(r, 2, f'=COUNTIF(B{kr}:T{kr},"✗")+COUNTIF(B{kr}:T{kr},"½")')
        ws.cell(r, 3, f'=IF(V{kr}="","✓ Semua selesai",LEFT(V{kr},LEN(V{kr})-2))')
        ws.cell(r, 11, f'=COUNTIF(B{lr}:T{lr},"✗")+COUNTIF(B{lr}:T{lr},"½")')
        ws.cell(r, 12, f'=IF(V{lr}="","✓ Semua selesai",LEFT(V{lr},LEN(V{lr})-2))')
        ws.merge_cells(start_row=r, start_column=3, end_row=r, end_column=10)
        ws.merge_cells(start_row=r, start_column=12, end_row=r, end_column=20)
        for c in (1, 2, 3, 11, 12):
            ws.cell(r, c).border = BOX
            ws.cell(r, c).font = font(bold=c in (1, 2, 11))
        for c in (2, 11):
            ws.cell(r, c).alignment = Alignment(horizontal="center", vertical="center")
        for c in (3, 12):
            ws.cell(r, c).alignment = Alignment(wrap_text=True, vertical="center")
        ws.row_dimensions[r].height = 30
    last = 5 + n
    for rng in (f"B6:B{last}", f"K6:K{last}"):
        ws.conditional_formatting.add(rng, CellIsRule(operator="greaterThan", formula=["0"], font=Font(name=FONT, bold=True, color=RED), fill=fill(RED_L)))
    for col in ("C", "L"):
        ws.conditional_formatting.add(f"{col}6:{col}{last}", FormulaRule(formula=[f'LEFT({col}6,1)="✓"'], font=Font(name=FONT, color=GREEN)))
    assert matrix(ws, top_k, "Jadual Yuran Kelab", "K") == k_first
    assert matrix(ws, top_l, "Jadual Yuran Lawatan", "L") == l_first
    ws.column_dimensions["V"].hidden = True
    widths(ws, (16,) + (8,) * 19)
    ws.freeze_panes = "B6"


# ---------------------------------------------------------------- Ringkasan
def build_ringkasan(ws):
    banner(ws, "Belanjawan Kelab Guru", "Ringkasan kewangan - dikemas kini automatik daripada helaian Transaksi", 12)
    tot_in, tot_out = f"SUM({T}G2:I{MAX})", f"SUM({T}J2:J{MAX})"
    first = "DATE(YEAR(TODAY()),MONTH(TODAY()),1)"

    def mon(cols, crit):
        return "+".join(f"SUMIFS({T}{c}2:{c}{MAX},{crit})" for c in cols)
    cur = f'{T}B2:B{MAX},">="&{first}'
    kpis = [("BAKI SEMASA", f"={tot_in}-{tot_out}", TEAL), ("JUMLAH MASUK", f"={tot_in}", GREEN),
            ("JUMLAH KELUAR", f"={tot_out}", RED), ("BULAN INI (BERSIH)", f"={mon('GHI', cur)}-({mon('J', cur)})", AMBER)]
    for i, (label, f, colr) in enumerate(kpis):
        c0 = 1 + i * 3
        for r in (4, 5):
            ws.merge_cells(start_row=r, start_column=c0, end_row=r, end_column=c0 + 1)
            for c in (c0, c0 + 1):
                ws.cell(r, c).fill = fill(GREY_L)
        a, b = ws.cell(4, c0, label), ws.cell(5, c0, f)
        a.font, b.font = font(size=9, bold=True, color=GREY), font(size=18, bold=True, color=colr)
        b.number_format = RM
        a.alignment = Alignment(indent=1, vertical="bottom")
        b.alignment = Alignment(indent=1, vertical="center", horizontal="left")
    ws.row_dimensions[4].height, ws.row_dimensions[5].height = 22, 34

    ws.cell(7, 1, "Pergerakan bulanan").font = font(size=13, bold=True, color=TEAL_D)
    header(ws, 8, ["Bulan", "Masuk", "Keluar", "Bersih", "Baki akhir"])
    for i in range(N_MONTHS):
        r = 9 + i
        ws.cell(r, 1, f"=EDATE(Tetapan!$B$22,{i})").number_format = "mmm yyyy"
        crit = f'{T}B2:B{MAX},">="&A{r},{T}B2:B{MAX},"<"&EDATE(A{r},1)'
        ws.cell(r, 2, f"={mon('GHI', crit)}")
        ws.cell(r, 3, f"={mon('J', crit)}")
        ws.cell(r, 4, f"=B{r}-C{r}")
        ws.cell(r, 5, f"=SUM(D$9:D{r})")
        for c in range(1, 6):
            cell = ws.cell(r, c)
            cell.font, cell.border = font(bold=(c == 5), color=GREEN if c == 2 else RED if c == 3 else INK), BOT
            if c > 1:
                cell.number_format = RM
    end = 8 + N_MONTHS
    ws.cell(end + 1, 1, "Jumlah").font = font(bold=True)
    for col in "BCD":
        c = ws[f"{col}{end + 1}"]
        c.value, c.number_format, c.font = f"=SUM({col}9:{col}{end})", RM, font(bold=True)
    for c in range(1, 6):
        ws.cell(end + 1, c).fill = fill(TEAL_L)
    ws.conditional_formatting.add(f"D9:D{end}", CellIsRule(operator="lessThan", formula=["0"], font=Font(name=FONT, color=RED, bold=True)))

    ws.cell(7, 7, "Perbelanjaan mengikut kategori").font = font(size=13, bold=True, color=TEAL_D)
    header(ws, 8, ["Kategori", "Jumlah", "%"], col=7)
    for i in range(7):
        r = 9 + i
        ws.cell(r, 7, f"=Tetapan!I{KAT_KELUAR_ROW + i}")
        ws.cell(r, 8, f"=SUMIFS({T}J2:J{MAX},{T}E2:E{MAX},G{r})").number_format = RM
        ws.cell(r, 9, f"=IF($C${end + 1}=0,0,H{r}/$C${end + 1})").number_format = "0%"
        for c in (7, 8, 9):
            ws.cell(r, c).font, ws.cell(r, c).border = font(), BOT

    ch = BarChart()
    ch.type, ch.grouping, ch.title, ch.height, ch.width = "col", "clustered", "Masuk vs Keluar", 8.5, 17
    ch.add_data(Reference(ws, min_col=2, max_col=3, min_row=8, max_row=end), titles_from_data=True)
    ch.set_categories(Reference(ws, min_col=1, min_row=9, max_row=end))
    ch.series[0].graphicalProperties.solidFill = GREEN
    ch.series[1].graphicalProperties.solidFill = RED
    ws.add_chart(ch, "G18")
    ln = LineChart()
    ln.title, ln.height, ln.width = "Baki akhir", 7.5, 17
    ln.add_data(Reference(ws, min_col=5, min_row=8, max_row=end), titles_from_data=True)
    ln.set_categories(Reference(ws, min_col=1, min_row=9, max_row=end))
    ln.series[0].graphicalProperties.line.solidFill = TEAL
    ln.series[0].graphicalProperties.line.width = 28000
    ws.add_chart(ln, "G35")
    widths(ws, (14, 15, 15, 15, 16, 3, 24, 15, 8))


def build(tx, out):
    wb = Workbook()
    ws_r = wb.active
    ws_r.title = "Ringkasan"
    ws_h, ws_t = wb.create_sheet("Harian"), wb.create_sheet("Transaksi")
    ws_y, ws_s = wb.create_sheet("Yuran Ahli"), wb.create_sheet("Tetapan")
    build_tetapan(ws_s)
    build_transaksi(ws_t, tx)
    build_harian(ws_h)
    build_yuran(ws_y)
    build_ringkasan(ws_r)
    for ws, colr in ((ws_r, TEAL), (ws_h, "8B5CF6"), (ws_t, "0EA5E9"), (ws_y, "F59E0B"), (ws_s, GREY)):
        ws.sheet_properties.tabColor = colr
    wb.save(out)


if __name__ == "__main__":
    build(json.load(open(sys.argv[1])), sys.argv[2])
