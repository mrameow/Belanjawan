"""Bina Belanjawan_Kelab_Guru_Baru.xlsx (minimalis; import ke Google Sheets).

Guna:  python3 -I tools/build_sheet.py <tx.json> <output.xlsx>

Tab: Dashboard | Harian | Yuran | Transaksi | Tetapan | Panduan
Transaksi (satu baris = satu bayaran / belanja):
  A Tarikh | B Butiran | C Ahli | D Kategori | E Bulan Mula | F Bulan Akhir |
  G Kelab | H Lawatan | I Masuk Lain | J Keluar | K Baki | L Nota | (M..R pembantu, tersembunyi)
"""
import datetime as dt
import json
import sys

from openpyxl import Workbook
from openpyxl.chart import BarChart, LineChart, Reference
from openpyxl.chart.shapes import GraphicalProperties
from openpyxl.drawing.text import CharacterProperties
from openpyxl.formatting.rule import CellIsRule, FormulaRule
from openpyxl.styles import Alignment, Border, Font, PatternFill, Side
from openpyxl.utils import get_column_letter as CL
from openpyxl.worksheet.datavalidation import DataValidation

# ---- palet minimalis: satu aksen (teal), merah/hijau hanya untuk angka
INK, MUTE, LINE, SOFT = "0F172A", "64748B", "E5E7EB", "F8FAFC"
TEAL, TEAL_L = "0F766E", "ECFDF5"
GREEN, GREEN_L, RED, RED_L, AMBER, AMBER_L = "15803D", "DCFCE7", "B91C1C", "FEE2E2", "B45309", "FEF3C7"
FONT = "Arial"
RM = '"RM"#,##0.00;[Red]-"RM"#,##0.00'
NUM = '#,##0.00;[Red]-#,##0.00;"–"'

MAXT = 2000            # baris Transaksi yang disediakan (baris 2..2001)
LAST = MAXT + 1
DAYS = 400             # bilangan hari berbeza yang disokong dalam Harian
LEDGER = 1200          # baris dalam Harian
N_MONTHS = 19          # Jun 2025 - Dis 2026
NSLOT = 20             # slot ahli dalam Tetapan
T = "Transaksi!"

AHLI = [  # nama, mula, tamat, catatan
    ("Amir", (2025, 6), None, ""), ("Asnawi", (2025, 6), None, ""), ("Bang Mie", (2025, 6), None, ""),
    ("Boss Khairil", (2025, 6), (2025, 9), "Pindah Sep 2025 (tab Conclusion)"), ("Boss Wan", (2025, 6), None, ""),
    ("Chang", (2025, 6), None, ""), ("Hadi", (2025, 6), None, ""), ("Hasnol", (2025, 6), None, ""),
    ("Joe", (2025, 6), None, ""), ("Khairiri", (2025, 6), None, ""),
    ("Kuchai", (2025, 6), None, "Lawatan = 0: refund RM400 Jan 2026. Semak"),
    ("Suhaimi", (2025, 6), (2025, 8), "Majlis perpisahan. Tamat anggaran"),
    ("Ustaz", (2025, 6), None, ""), ("Wafi", (2025, 6), (2026, 3), "Perpisahan Mac 2026. Lawatan = 0, semak"),
    ("Aiman", (2025, 12), None, "Masuk bulan 12 (tab Conclusion)"),
    ("Yusuf", (2026, 9), None, "Ahli baru dalam tab Yuran Kelab lama. Semak mula & kadar"),
    ("Fairuz", (2026, 10), None, "Ahli baru dalam tab Yuran Kelab lama. Semak mula & kadar"),
]
NO_LAWATAN = {"Kuchai", "Wafi"}
KATEGORI = ["Yuran", "Majlis", "Sumbangan", "Lain-lain",                       # masuk
            "Duit Pantri", "Lawatan / Trip", "Majlis & Hadiah", "Pakaian",      # keluar
            "Langganan / Bahan", "Refund", "Pelbagai"]
KAT_KELUAR_ROW = 9  # Tetapan!I9..I15


def fill(c):
    return PatternFill("solid", start_color=c, end_color=c)


def font(**kw):
    kw.setdefault("name", FONT)
    kw.setdefault("size", 10)
    kw.setdefault("color", INK)
    return Font(**kw)


hair = Side(style="thin", color=LINE)
BOT = Border(bottom=hair)


def title(ws, text, sub, col=1):
    ws.sheet_view.showGridLines = False
    ws.cell(1, col, text).font = font(size=18, bold=True)
    ws.cell(2, col, sub).font = font(size=9, color=MUTE)
    ws.row_dimensions[1].height = 30


def head(ws, row, labels, col=1, align="center", height=22):
    for i, t in enumerate(labels):
        c = ws.cell(row, col + i, t)
        c.font = font(size=9, bold=True, color=MUTE)
        c.alignment = Alignment(horizontal=align if i else "left", vertical="center", wrap_text=True)
        c.border = Border(bottom=Side(style="medium", color=INK))
    ws.row_dimensions[row].height = height


def widths(ws, ws_widths, start=1):
    for i, w in enumerate(ws_widths):
        ws.column_dimensions[CL(start + i)].width = w


def cf(ws, rng, formula=None, op=None, value=None, **style):
    st = dict(font=style.get("font"), fill=style.get("fill"))
    if formula:
        ws.conditional_formatting.add(rng, FormulaRule(formula=[formula], **{k: v for k, v in st.items() if v}))
    else:
        ws.conditional_formatting.add(rng, CellIsRule(operator=op, formula=[value], **{k: v for k, v in st.items() if v}))


# ===================================================================== Tetapan
def build_tetapan(ws):
    title(ws, "Tetapan", "Senarai ahli, kadar dan kategori. Tab lain membaca helaian ini. Sel kuning boleh diubah.")
    head(ws, 4, ["Ahli", "Mula", "Tamat", "Kelab / bulan", "Lawatan / bulan", "Catatan"], align="center")
    for i in range(NSLOT):
        r = 5 + i
        if i < len(AHLI):
            n, mula, tamat, note = AHLI[i]
            vals = [n, dt.date(*mula, 1), dt.date(*tamat, 1) if tamat else None, 35, 0 if n in NO_LAWATAN else 50, note]
        else:
            vals = [None] * 6
        for c, v in enumerate(vals, 1):
            cell = ws.cell(r, c, v)
            cell.font, cell.border = font(color=MUTE if c == 6 else INK, size=9 if c == 6 else 10), BOT
            if c in (2, 3):
                cell.number_format, cell.alignment = "mmm yyyy", Alignment(horizontal="center")
            if c in (4, 5):
                cell.number_format = '"RM"#,##0'
                cell.alignment = Alignment(horizontal="center")
    ws.cell(27, 1, "Lawatan dikutip sehingga").font = font(bold=True)
    ws.cell(28, 1, "Mula rekod").font = font(bold=True)
    for r, d in ((27, dt.date(2026, 2, 1)), (28, dt.date(2025, 6, 1))):
        c = ws.cell(r, 2, d)
        c.number_format, c.fill, c.font, c.alignment = "mmm yyyy", fill("FEF9C3"), font(), Alignment(horizontal="center")
    ws.cell(30, 1, "Tamat kosong = masih aktif. Kadar 0 = dikecualikan. Tambah ahli baru pada baris kosong (sehingga 20 ahli).").font = font(size=9, italic=True, color=MUTE)
    head(ws, 4, ["Kategori"], col=8, align="left")
    for i, v in enumerate(KATEGORI):
        ws.cell(5 + i, 8, v).font = font()
        ws.cell(5 + i, 8).border = BOT
    ws.cell(5, 9, "masuk").font = font(size=8, color=MUTE)
    ws.cell(KAT_KELUAR_ROW, 9, "keluar").font = font(size=8, color=MUTE)
    for r in range(5, 5 + NSLOT):
        for c in (2, 3):
            if ws.cell(r, c).value is not None or r < 5 + NSLOT:
                ws.cell(r, c).fill = fill("FEF9C3") if r < 5 + len(AHLI) else PatternFill()
    widths(ws, (22, 11, 11, 14, 16, 52, 3, 20, 8))
    ws.freeze_panes = "A5"


# =================================================================== Transaksi
def build_transaksi(ws, tx):
    ws.sheet_view.showGridLines = False
    head(ws, 1, ["Tarikh", "Butiran", "Ahli", "Kategori", "Bulan mula", "Bulan akhir", "Kelab", "Lawatan",
                 "Masuk lain", "Keluar", "Baki", "Nota"], height=26)
    for c in (1, 2, 3, 4, 12):
        ws.cell(1, c).alignment = Alignment(horizontal="left", vertical="center")
    for i, x in enumerate(tx):
        mula = dt.date.fromisoformat(x["mula"] + "-01") if x["mula"] else None
        akhir = dt.date.fromisoformat(x["akhir"] + "-01") if x["akhir"] and x["akhir"] != x["mula"] else None
        vals = [dt.date.fromisoformat(x["tarikh"]), x["butiran"], x["ahli"] or None, x["kategori"], mula, akhir,
                x["kelab"] or None, x["lawatan"] or None, x["lain"] or None, x["keluar"] or None]
        for c, v in enumerate(vals, 1):
            ws.cell(i + 2, c, v)
        if x.get("nota"):
            ws.cell(i + 2, 12, x["nota"])
    for r in range(2, LAST + 1):
        ws.cell(r, 11, f'=IF(A{r}="","",SUM(G$2:I{r})-SUM(J$2:J{r}))')
        # pembantu: M hari#, N mula(angka), O akhir(angka), P bil.bulan, Q kelab/bln, R lawatan/bln
        ws.cell(r, 13, f'=IF(A{r}="","",IF(A{r}=A{r - 1},M{r - 1},N(M{r - 1})+1))')
        ws.cell(r, 14, f'=IF(E{r}="",99999,E{r})')
        ws.cell(r, 15, f'=IF(E{r}="",0,IF(F{r}="",E{r},F{r}))')
        ws.cell(r, 16, f'=IF(E{r}="",1,MAX(1,(YEAR(O{r})-YEAR(N{r}))*12+MONTH(O{r})-MONTH(N{r})+1))')
        ws.cell(r, 17, f'=IF(G{r}="",0,G{r}/P{r})')
        ws.cell(r, 18, f'=IF(H{r}="",0,H{r}/P{r})')
        for c in range(1, 13):
            cell = ws.cell(r, c)
            cell.font, cell.border = font(color=MUTE if c in (11, 12) else INK, size=9 if c == 12 else 10), BOT
        ws.cell(r, 1).number_format = "dd mmm yyyy"
        ws.cell(r, 1).alignment = Alignment(horizontal="left")
        for c in (5, 6):
            ws.cell(r, c).number_format = "mmm yyyy"
            ws.cell(r, c).alignment = Alignment(horizontal="center")
        for c in range(7, 12):
            ws.cell(r, c).number_format = NUM
    # hari genap berlatar lembut -> setiap hari nampak terasing
    cf(ws, f"A2:L{LAST}", formula='AND($M2<>"",ISEVEN($M2))', fill=fill(SOFT))
    cf(ws, f"A2:A{LAST}", formula='AND($A2<>"",$A2=$A1)', font=Font(name=FONT, color="CBD5E1"))
    cf(ws, f"A2:A{LAST}", formula='AND($A2<>"",$A2<>$A1)', font=Font(name=FONT, bold=True))
    cf(ws, f"G2:I{LAST}", op="greaterThan", value="0", font=Font(name=FONT, color=GREEN))
    cf(ws, f"J2:J{LAST}", op="greaterThan", value="0", font=Font(name=FONT, color=RED))
    cf(ws, f"E2:F{LAST}", formula='AND($D2="Yuran",$E2="")', fill=fill(AMBER_L))   # yuran tanpa bulan
    cf(ws, f"L2:L{LAST}", formula='$L2<>""', font=Font(name=FONT, color=AMBER, bold=True))
    for col, src in (("D", "$H$5:$H$15"), ("C", "$A$5:$A$24")):
        dv = DataValidation(type="list", formula1=f"=Tetapan!{src}", allow_blank=True, showErrorMessage=False)
        ws.add_data_validation(dv)
        dv.add(f"{col}2:{col}{LAST}")
    widths(ws, (13, 40, 14, 16, 11, 11, 11, 11, 11, 11, 12, 22))
    for c in range(13, 19):
        ws.column_dimensions[CL(c)].hidden = True
    ws.freeze_panes = "A2"
    ws.auto_filter.ref = f"A1:L{LAST}"


# ====================================================================== Harian
def build_harian(ws):
    title(ws, "Harian", "Setiap hari: semua wang masuk & keluar pada hari itu. Terbaru di atas. Auto daripada Transaksi.")
    # --- pembantu hari (kolum N..Q, baris 3..): N bil, O hari#, P saiz blok, Q mula baris
    L = f"{T}$M$2:$M${LAST}"
    ws["N1"], ws["O1"] = "maksimum hari", f"=MAX({L})"
    ws["N2"], ws["O2"] = "jumlah baris", f"=SUM(P3:P{2 + DAYS})"
    for j in range(1, DAYS + 1):
        r = 2 + j
        ws.cell(r, 14, j)
        ws.cell(r, 15, f'=IF(N{r}<=$O$1,$O$1-N{r}+1,"")')
        ws.cell(r, 16, f'=IF(O{r}="",0,COUNTIF({L},O{r})+1)')
        ws.cell(r, 17, f"=1+SUM(P$2:P{r - 1})")
    head(ws, 3, ["", "Butiran", "Ahli", "Kategori", "Masuk", "Keluar", "Baki hujung hari"], height=24)
    ws.cell(3, 1).value = None
    jr, tr = f"$P$3:$P${2 + DAYS}", None
    for k in range(1, LEDGER + 1):
        r = 3 + k
        # I: j (indeks hari), J: jenis H/E, K: indeks baris Transaksi, S: nombor hari, T: tarikh
        ws.cell(r, 9, f'=IF({k}>$O$2,"",MATCH({k},$Q$3:$Q${2 + DAYS},1))')
        ws.cell(r, 10, f'=IF(I{r}="","",IF({k}=INDEX($Q$3:$Q${2 + DAYS},I{r}),"H","E"))')
        ws.cell(r, 19, f'=IF(I{r}="","",INDEX($O$3:$O${2 + DAYS},I{r}))')
        ws.cell(r, 11, f'=IF(J{r}<>"E","",MATCH(S{r},{L},0)+{k}-INDEX($Q$3:$Q${2 + DAYS},I{r})-1)')
        ws.cell(r, 20, f'=IF(J{r}<>"H","",INDEX({T}$A$2:$A${LAST},MATCH(S{r},{L},0)))')

        def tx(col):
            return f"INDEX({T}${col}$2:${col}${LAST},K{r})"

        def day(cols, op="="):
            return "+".join(f'SUMIFS({T}${c}$2:${c}${LAST},{L},"{op}"&S{r})' for c in cols)
        wd = f'CHOOSE(WEEKDAY(T{r},2),"Isnin","Selasa","Rabu","Khamis","Jumaat","Sabtu","Ahad")'
        mn = f'CHOOSE(MONTH(T{r}),"Jan","Feb","Mac","Apr","Mei","Jun","Jul","Ogo","Sep","Okt","Nov","Dis")'
        ws.cell(r, 1, f'=IF(J{r}="E",IF({tx("J")}>0,"↑","↓"),"")')
        ws.cell(r, 2, f'=IF(J{r}="H",{wd}&", "&DAY(T{r})&" "&{mn}&" "&YEAR(T{r}),IF(J{r}="E",IF({tx("B")}="",{tx("C")}&" · "&{tx("D")},{tx("B")}),""))')
        ws.cell(r, 3, f'=IF(J{r}="E",{tx("C")}&"","")')
        ws.cell(r, 4, f'=IF(J{r}="E",{tx("D")}&"","")')
        ws.cell(r, 5, f'=IF(J{r}="H",{day("GHI")},IF(J{r}="E",IF({tx("G")}+{tx("H")}+{tx("I")}=0,"",{tx("G")}+{tx("H")}+{tx("I")}),""))')
        ws.cell(r, 6, f'=IF(J{r}="H",{day("J")},IF(J{r}="E",IF({tx("J")}=0,"",{tx("J")}),""))')
        ws.cell(r, 7, f'=IF(J{r}="H",({day("GHI", "<=")})-({day("J", "<=")}),"")')
        for c in range(1, 8):
            cell = ws.cell(r, c)
            cell.font = font()
            if c >= 5:
                cell.number_format = NUM
        ws.cell(r, 1).alignment = Alignment(horizontal="center")
    last = 3 + LEDGER
    cf(ws, f"A4:G{last}", formula='$J4="H"', fill=fill(TEAL_L), font=Font(name=FONT, bold=True, color=TEAL))
    cf(ws, f"E4:E{last}", formula='$J4="E"', font=Font(name=FONT, color=GREEN))
    cf(ws, f"F4:F{last}", formula='$J4="E"', font=Font(name=FONT, color=RED))
    cf(ws, f"A4:A{last}", formula='$A4="↑"', font=Font(name=FONT, color=RED, bold=True))
    cf(ws, f"A4:A{last}", formula='$A4="↓"', font=Font(name=FONT, color=GREEN, bold=True))
    cf(ws, f"C4:D{last}", formula='$J4="E"', font=Font(name=FONT, color=MUTE))
    widths(ws, (4, 44, 14, 16, 13, 13, 17, 2))
    for c in range(9, 21):
        ws.column_dimensions[CL(c)].hidden = True
    ws.freeze_panes = "A4"


# ======================================================================= Yuran
def grid(ws, top, label, kind):
    """kind K (kelab; lajur Q) atau L (lawatan; lajur R). Mengembalikan baris data pertama."""
    q, rate = ("Q", "D") if kind == "K" else ("R", "E")
    ws.cell(top, 1, label).font = font(size=12, bold=True)
    h = top + 1
    head(ws, h, ["Ahli"], height=24)
    for i in range(N_MONTHS):
        c = ws.cell(h, 2 + i, "=Tetapan!$B$28" if i == 0 else f"=EDATE({CL(1 + i)}{h},1)")
        c.number_format = "mmm yy"
        c.font = font(size=8, bold=True, color=MUTE)
        c.alignment = Alignment(horizontal="center", vertical="center")
        c.border = Border(bottom=Side(style="medium", color=INK))
    first = h + 1
    N, O, C, Qr = (f"{T}${c}$2:${c}${LAST}" for c in ("N", "O", "C", q))
    for k in range(NSLOT):
        r, s = first + k, 5 + k
        ws.cell(r, 1, f'=IF(Tetapan!A{s}="","",Tetapan!A{s})').font = font(bold=True)
        ws.cell(r, 1).border = BOT
        for i in range(N_MONTHS):
            m = f"{CL(2 + i)}${h}"
            who = f"({C}=$A{r})*({N}<={m})*({O}>={m})"
            out = (f'OR({m}<Tetapan!$B{s},AND(Tetapan!$C{s}<>"",{m}>Tetapan!$C{s}),Tetapan!${rate}{s}=0'
                   + (f",{m}>Tetapan!$B$27)" if kind == "L" else ")"))
            f = (f'=IF($A{r}="","",IF({out},"–",IF(SUMPRODUCT({who}*({Qr}>=Tetapan!${rate}{s}-0.01))>0,"✓",'
                 f'IF(SUMPRODUCT({who}*({Qr}>0))>0,"½",IF({m}>TODAY(),"·","✗")))))')
            cell = ws.cell(r, 2 + i, f)
            cell.alignment, cell.border, cell.font = Alignment(horizontal="center"), BOT, font(bold=True, size=10)
        parts = "&".join(
            f'IF({CL(2 + i)}{r}="✗",TEXT({CL(2 + i)}${h},"mmm yy")&", ",'
            f'IF({CL(2 + i)}{r}="½",TEXT({CL(2 + i)}${h},"mmm yy")&" (separa), ",""))' for i in range(N_MONTHS))
        ws.cell(r, 23, f"={parts}").font = font(size=8, color=MUTE)
    rng = f"B{first}:T{first + NSLOT - 1}"
    cf(ws, rng, op="equal", value='"✓"', font=Font(name=FONT, bold=True, color=GREEN), fill=fill(GREEN_L))
    cf(ws, rng, op="equal", value='"✗"', font=Font(name=FONT, bold=True, color=RED), fill=fill(RED_L))
    cf(ws, rng, op="equal", value='"½"', font=Font(name=FONT, bold=True, color=AMBER), fill=fill(AMBER_L))
    cf(ws, rng, op="equal", value='"–"', font=Font(name=FONT, color="CBD5E1"))
    cf(ws, rng, op="equal", value='"·"', font=Font(name=FONT, color="CBD5E1"))
    return first


def build_yuran(ws):
    title(ws, "Yuran", "✓ dibayar   ½ separa   ✗ belum bayar   · akan datang   – di luar tempoh / dikecualikan.   Auto daripada Transaksi (lajur Bulan mula/akhir).")
    head(ws, 4, ["Ahli", "Kelab belum", "Bulan Kelab yang belum dibayar", "", "", "", "", "", "", "",
                 "Lawatan belum", "Bulan Lawatan yang belum dibayar", "", "", "", "", "", "", "", ""], height=30)
    ws.merge_cells("C4:J4")
    ws.merge_cells("L4:T4")
    ws["C4"].alignment = ws["L4"].alignment = Alignment(horizontal="left", vertical="center")
    k_top = 5 + NSLOT + 2
    k_first, l_top = k_top + 2, k_top + 2 + NSLOT + 2
    l_first = l_top + 2
    for k in range(NSLOT):
        r, kr, lr = 5 + k, k_first + k, l_first + k
        ws.cell(r, 1, f'=IF(Tetapan!A{5 + k}="","",Tetapan!A{5 + k})')
        ws.cell(r, 2, f'=IF($A{r}="","",COUNTIF(B{kr}:T{kr},"✗")+COUNTIF(B{kr}:T{kr},"½"))')
        ws.cell(r, 3, f'=IF($A{r}="","",IF(W{kr}="","✓ semua selesai",LEFT(W{kr},LEN(W{kr})-2)))')
        ws.cell(r, 11, f'=IF($A{r}="","",COUNTIF(B{lr}:T{lr},"✗")+COUNTIF(B{lr}:T{lr},"½"))')
        ws.cell(r, 12, f'=IF($A{r}="","",IF(W{lr}="","✓ semua selesai",LEFT(W{lr},LEN(W{lr})-2)))')
        ws.cell(r, 24, f'=IF($A{r}="",0,COUNTIF(B{kr}:T{kr},"✗")*Tetapan!D{5 + k}+COUNTIF(B{lr}:T{lr},"✗")*Tetapan!E{5 + k})')
        ws.merge_cells(start_row=r, start_column=3, end_row=r, end_column=10)
        ws.merge_cells(start_row=r, start_column=12, end_row=r, end_column=20)
        for c in range(1, 21):
            ws.cell(r, c).border = BOT
        ws.cell(r, 1).font = font(bold=True)
        ws.cell(r, 1).alignment = Alignment(vertical="center")
        for c in (2, 11):
            ws.cell(r, c).alignment = Alignment(horizontal="center", vertical="center")
            ws.cell(r, c).font = font(bold=True)
        for c in (3, 12):
            ws.cell(r, c).alignment = Alignment(wrap_text=True, vertical="center")
            ws.cell(r, c).font = font(size=9, color=MUTE)
        ws.row_dimensions[r].height = 26
    last = 4 + NSLOT
    for col in ("B", "K"):
        cf(ws, f"{col}5:{col}{last}", op="greaterThan", value="0", font=Font(name=FONT, bold=True, color=RED), fill=fill(RED_L))
    for col in ("C", "L"):
        cf(ws, f"{col}5:{col}{last}", formula=f'LEFT({col}5,1)="✓"', font=Font(name=FONT, size=9, color=GREEN))
    assert grid(ws, k_top, "Kelab", "K") == k_first
    assert grid(ws, l_top, "Lawatan", "L") == l_first
    ws.column_dimensions["W"].hidden = True
    ws.column_dimensions["X"].hidden = True
    widths(ws, (16,) + (7.6,) * 19)
    ws.freeze_panes = "B5"


# =================================================================== Dashboard
def build_dashboard(ws):
    title(ws, "Belanjawan Kelab Guru", "", col=2)
    ws["B2"] = '="Dikemas kini "&TEXT(TODAY(),"d mmm yyyy")&"  ·  data daripada tab Transaksi"'
    ws["B2"].font = font(size=9, color=MUTE)
    widths(ws, (2,) + (10.5,) * 12 + (2,))

    def card(row, col, span, label, formula, color=INK, fmt=RM, size=18, caption=None):
        for r in (row, row + 1) + ((row + 2,) if caption else ()):
            ws.merge_cells(start_row=r, start_column=col, end_row=r, end_column=col + span - 1)
            for c in range(col, col + span):
                ws.cell(r, c).fill = fill(SOFT)
        a, b = ws.cell(row, col, label), ws.cell(row + 1, col, formula)
        a.font, b.font = font(size=8, bold=True, color=MUTE), font(size=size, bold=True, color=color)
        b.number_format = fmt
        for c_ in (a, b):
            c_.alignment = Alignment(horizontal="left", vertical="center", indent=1)
        if caption:
            d = ws.cell(row + 2, col, caption)
            d.font, d.alignment = font(size=8, color=MUTE), Alignment(horizontal="left", vertical="top", indent=1, wrap_text=True)

    A, G, H, I, J, Dt = (f"{T}${c}$2:${c}${LAST}" for c in "AGHIJA")
    tin, tout = f"SUM({T}$G$2:$I${LAST})", f"SUM({J})"
    card(4, 2, 3, "BAKI SEMASA", f"={tin}-{tout}", TEAL, size=20)
    card(4, 5, 3, "JUMLAH MASUK", f"={tin}", GREEN, size=20)
    card(4, 8, 3, "JUMLAH KELUAR", f"={tout}", RED, size=20)
    card(4, 11, 3, "YURAN BELUM BAYAR", "=SUM(Yuran!$X$5:$X$24)", AMBER, size=20, caption="anggaran, ikut jadual Yuran")
    for r, h in ((4, 20), (5, 34), (6, 16)):
        ws.row_dimensions[r].height = h

    # ---- pemilih bulan
    ws["B8"] = "Lihat bulan"
    ws["B8"].font = font(size=11, bold=True)
    ws.merge_cells("C8:D8")
    ws["C8"] = "=DATE(YEAR(TODAY()),MONTH(TODAY()),1)"
    ws["C8"].number_format = "mmmm yyyy"
    ws["C8"].font = font(bold=True, color=TEAL)
    ws["C8"].fill = fill("FEF9C3")
    ws["C8"].alignment = Alignment(horizontal="center")
    ws["E8"] = "◂ pilih bulan daripada senarai"
    ws["E8"].font = font(size=8, color=MUTE)
    mt = 52  # baris pertama jadual bulanan
    dv = DataValidation(type="list", formula1=f"=$B${mt}:$B${mt + N_MONTHS - 1}", allow_blank=False)
    ws.add_data_validation(dv)
    dv.add("C8")
    sel = "$C$8"
    nxt = f"EDATE({sel},1)"
    mcrit = f'{A},">="&{sel},{A},"<"&{nxt}'
    msum = lambda cols: "+".join(f"SUMIFS({T}${c}$2:${c}${LAST},{mcrit})" for c in cols)
    ws.row_dimensions[10].height, ws.row_dimensions[11].height, ws.row_dimensions[12].height = 18, 26, 26
    yk_h = 5 + NSLOT + 3          # baris tajuk bulan jadual Kelab dalam Yuran
    kc = f"INDEX(Yuran!$B${yk_h + 1}:$T${yk_h + NSLOT},0,MATCH({sel},Yuran!$B${yk_h}:$T${yk_h},0))"
    maxv = f"SUMPRODUCT(MAX(({J})*({A}>={sel})*({A}<{nxt})))"
    row_of = f"SUMPRODUCT(MAX(({J}={maxv})*({A}>={sel})*({A}<{nxt})*ROW({J})))"
    card(10, 2, 2, "MASUK", f"={msum('GHI')}", GREEN, size=13)
    card(10, 4, 2, "KELUAR", f"={msum('J')}", RED, size=13)
    card(10, 6, 2, "BERSIH", f"={msum('GHI')}-({msum('J')})", INK, size=13)
    card(10, 8, 2, "BIL. REKOD", f'=COUNTIFS({mcrit})', INK, fmt="0", size=13)
    card(10, 10, 2, "KELAB DIBAYAR", f'=COUNTIF({kc},"✓")&" / "&(COUNTIF({kc},"✓")+COUNTIF({kc},"✗")+COUNTIF({kc},"½"))&" ahli"', TEAL, fmt="@", size=13)
    card(10, 12, 2, "BELANJA TERBESAR", f"={maxv}", RED, size=13,
         caption=f'=IF({maxv}=0,"-",INDEX({T}$B$1:$B${LAST},{row_of}))')
    ws.row_dimensions[12].height = 26

    # ---- carta
    ws["B14"] = "Graf"
    ws["B14"].font = font(size=11, bold=True)
    mend = mt + N_MONTHS - 1
    cats = Reference(ws, min_col=2, min_row=mt, max_row=mend)

    def style(ch, ttl, w=12.2, h=7.2):
        ch.title, ch.width, ch.height, ch.legend.position = ttl, w, h, "b"
        ch.title.tx.rich.p[0].r[0].rPr = CharacterProperties(sz=1100, b=True, solidFill=INK)
        ch.y_axis.number_format = "#,##0"
        ch.y_axis.majorGridlines = None
        ch.y_axis.delete = False
        ch.x_axis.delete = False
        ch.x_axis.number_format = "mmm yy"
        return ch
    c1 = style(BarChart(), "Masuk & keluar setiap bulan")
    c1.type, c1.grouping = "col", "clustered"
    c1.add_data(Reference(ws, min_col=3, max_col=4, min_row=mt - 1, max_row=mend), titles_from_data=True)
    c1.set_categories(cats)
    c1.series[0].graphicalProperties = GraphicalProperties(solidFill="16A34A")
    c1.series[1].graphicalProperties = GraphicalProperties(solidFill="DC2626")
    ws.add_chart(c1, "B15")
    c2 = style(LineChart(), "Baki akhir bulan")
    c2.add_data(Reference(ws, min_col=6, min_row=mt - 1, max_row=mend), titles_from_data=True)
    c2.set_categories(cats)
    c2.series[0].graphicalProperties.line.solidFill = TEAL
    c2.series[0].graphicalProperties.line.width = 28000
    c2.series[0].smooth = False
    c2.legend = None
    ws.add_chart(c2, "H15")
    c3 = style(BarChart(), "Sumber wang masuk")
    c3.type, c3.grouping, c3.overlap = "col", "stacked", 100
    c3.add_data(Reference(ws, min_col=7, max_col=9, min_row=mt - 1, max_row=mend), titles_from_data=True)
    c3.set_categories(cats)
    for s_, col in zip(c3.series, ("0F766E", "38BDF8", "A3A3A3")):
        s_.graphicalProperties = GraphicalProperties(solidFill=col)
    ws.add_chart(c3, "B31")
    ck = BarChart()
    ck.type = "bar"
    style(ck, "Perbelanjaan mengikut kategori")
    ck.add_data(Reference(ws, min_col=3, min_row=mt + N_MONTHS + 5, max_row=mt + N_MONTHS + 11), titles_from_data=False)
    ck.set_categories(Reference(ws, min_col=2, min_row=mt + N_MONTHS + 5, max_row=mt + N_MONTHS + 11))
    ck.series[0].graphicalProperties = GraphicalProperties(solidFill="DC2626")
    ck.legend = None
    ws.add_chart(ck, "H31")

    # ---- jadual bulanan
    ws.cell(mt - 3, 2, "Pergerakan bulanan").font = font(size=11, bold=True)
    head(ws, mt - 1, ["Bulan", "Masuk", "Keluar", "Bersih", "Baki akhir", "Kelab", "Lawatan", "Lain"], col=2)
    for i in range(N_MONTHS):
        r = mt + i
        ws.cell(r, 2, f"=EDATE(Tetapan!$B$28,{i})").number_format = "mmm yyyy"
        crit = f'{A},">="&$B{r},{A},"<"&EDATE($B{r},1)'
        sm = lambda cols: "+".join(f"SUMIFS({T}${c}$2:${c}${LAST},{crit})" for c in cols)
        ws.cell(r, 3, f"={sm('GHI')}")
        ws.cell(r, 4, f"={sm('J')}")
        ws.cell(r, 5, f"=C{r}-D{r}")
        ws.cell(r, 6, f"=SUM(E${mt}:E{r})")
        ws.cell(r, 7, f"={sm('G')}")
        ws.cell(r, 8, f"={sm('H')}")
        ws.cell(r, 9, f"={sm('I')}")
        for c in range(2, 10):
            cell = ws.cell(r, c)
            cell.font, cell.border = font(bold=(c == 6), color=GREEN if c == 3 else RED if c == 4 else INK), BOT
            if c > 2:
                cell.number_format = NUM
        ws.cell(r, 2).alignment = Alignment(horizontal="left")
    tr = mt + N_MONTHS
    ws.cell(tr, 2, "Jumlah").font = font(bold=True)
    for c in (3, 4, 5, 7, 8, 9):
        cell = ws.cell(tr, c, f"=SUM({CL(c)}{mt}:{CL(c)}{tr - 1})")
        cell.number_format, cell.font = NUM, font(bold=True)
    for c in range(2, 10):
        ws.cell(tr, c).fill = fill(SOFT)
    cf(ws, f"E{mt}:E{tr - 1}", op="lessThan", value="0", font=Font(name=FONT, color=RED, bold=True))
    cf(ws, f"B{mt}:I{tr - 1}", formula=f"$B{mt}={sel}", fill=fill(TEAL_L))   # sorot bulan dipilih

    # ---- kategori perbelanjaan
    ct = tr + 3
    ws.cell(ct, 2, "Perbelanjaan mengikut kategori").font = font(size=11, bold=True)
    head(ws, ct + 1, ["Kategori", "Jumlah", "%", ""], col=2)
    ws.merge_cells(start_row=ct + 1, start_column=5, end_row=ct + 1, end_column=9)
    for i in range(7):
        r = ct + 2 + i
        ws.cell(r, 2, f"=Tetapan!H{KAT_KELUAR_ROW + i}")
        ws.cell(r, 3, f"=SUMIFS({J},{T}$D$2:$D${LAST},B{r})").number_format = NUM
        ws.cell(r, 4, f"=IF(SUM(C${ct + 2}:C${ct + 8})=0,0,C{r}/SUM(C${ct + 2}:C${ct + 8}))").number_format = "0%"
        ws.cell(r, 5, f'=REPT("█",ROUND(D{r}*40,0))')
        for c in range(2, 6):
            ws.cell(r, c).font, ws.cell(r, c).border = font(color=RED if c == 5 else INK, size=8 if c == 5 else 10), BOT
    assert ct + 2 == mt + N_MONTHS + 5, (ct, mt)
    ws.sheet_view.zoomScale = 100


# ==================================================================== Panduan
def build_panduan(ws):
    title(ws, "Panduan ringkas", "Lima perkara yang perlu diingat.")
    steps = [
        ("1. Tambah rekod", "Hanya di tab Transaksi: isi baris kosong paling bawah (Tarikh, Butiran, Ahli, Kategori, jumlah). Selebihnya auto."),
        ("2. Wang masuk / keluar", "Yuran → lajur Kelab / Lawatan. Majlis, sumbangan → Masuk lain. Perbelanjaan → Keluar."),
        ("3. Bayaran yuran", "Isi Bulan mula dan Bulan akhir (cth. Mac 2026 hingga Apr 2026). Satu bulan sahaja: kosongkan Bulan akhir. Bulan tidak berturutan: guna dua baris."),
        ("4. Bayaran separa", "Satu baris satu bulan dengan jumlah kurang daripada kadar. Tab Yuran menunjukkan ½."),
        ("5. Susunan", "Pastikan baris tersusun ikut tarikh (Data → Sort range A→Z pada lajur Tarikh) supaya tab Harian tepat."),
        ("Sel amber", "Lajur Bulan mula berwarna amber = bayaran yuran tanpa bulan. Isi bulan supaya tab Yuran mengira bayaran itu."),
        ("Tetapan", "Tambah ahli, tukar kadar atau tarikh Mula/Tamat di tab Tetapan."),
    ]
    for i, (a, b) in enumerate(steps):
        r = 4 + i * 2
        ws.cell(r, 1, a).font = font(bold=True, color=TEAL)
        c = ws.cell(r, 2, b)
        c.font, c.alignment = font(), Alignment(wrap_text=True, vertical="top")
        ws.row_dimensions[r].height = 32
    widths(ws, (22, 100))


def build(tx, out):
    wb = Workbook()
    ws_d = wb.active
    ws_d.title = "Dashboard"
    ws_h, ws_y = wb.create_sheet("Harian"), wb.create_sheet("Yuran")
    ws_t, ws_s, ws_p = wb.create_sheet("Transaksi"), wb.create_sheet("Tetapan"), wb.create_sheet("Panduan")
    build_tetapan(ws_s)
    build_transaksi(ws_t, tx)
    build_harian(ws_h)
    build_yuran(ws_y)
    build_dashboard(ws_d)
    build_panduan(ws_p)
    for ws, colr in ((ws_d, TEAL), (ws_h, "8B5CF6"), (ws_y, "F59E0B"), (ws_t, "0EA5E9"), (ws_s, MUTE), (ws_p, "CBD5E1")):
        ws.sheet_properties.tabColor = colr
    wb.save(out)


if __name__ == "__main__":
    build(json.load(open(sys.argv[1])), sys.argv[2])
