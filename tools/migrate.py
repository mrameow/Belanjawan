"""Tukar helaian lama (2025 + 2026) kepada senarai transaksi yang bersih.

Guna:  python3 -I tools/migrate.py <Belanjawan_Kelab_Guru.xlsx> <output.json>

Satu baris sumber = satu transaksi: Tarikh, Ahli, Butiran, Kategori,
Bulan (bulan yuran yang dibayar, cth "2026-03, 2026-04"), Kelab, Lawatan, Lain, Keluar.
Nombor dalam kurungan, cth "Amir (3, 4)", ialah BULAN yang dibayar.
  "12[30]"            -> bulan 12 dibayar separa RM30   (token "2025-12:30")
  "setengah bulan 12" -> separuh kadar bulan 12
"""
import datetime as dt
import json
import re
import sys

import openpyxl

AHLI = ["Amir", "Asnawi", "Bang Mie", "Boss Khairil", "Boss Wan", "Chang", "Hadi",
        "Hasnol", "Joe", "Khairiri", "Kuchai", "Suhaimi", "Ustaz", "Wafi", "Aiman"]
ALIAS = {"Boss": "Boss Khairil"}  # helaian 2025 tulis "Boss (6)"; sepadan tepat dengan tab Yuran Kelab (Boss Khairil)
RATE = {"kelab": 35, "lawatan": 50}
START = (2025, 6)  # bulan pertama rekod


def to_date(v):
    """Excel menukar 'dd/mm' kepada mm/dd bila dd <= 12 -> tukar balik."""
    if isinstance(v, dt.datetime):
        return dt.date(v.year, v.day, v.month)
    d, m, y = v.strip().split("/")
    return dt.date(int(y), int(m), int(d))


def find_ahli(text):
    t = text.strip()
    m = re.match(r"^([A-Za-z ]+?)\s*(\(|$)", t)
    head = ALIAS.get(m.group(1).strip(), m.group(1).strip()) if m else t
    if head in AHLI:
        return head
    m = re.search(r"-\s*([A-Za-z ]+?)\s*(\(|$)", t)  # "Majlis KKGS - Hasnol (Lamb Chop)"
    if m:
        tail = ALIAS.get(m.group(1).strip(), m.group(1).strip())
        if tail in AHLI:
            return tail
    return ""


def kategori_masuk(text):
    low = text.lower()
    if low.startswith("majlis"):
        return "Majlis"
    if "sumbangan" in low:
        return "Sumbangan"
    return "Lain-lain"


def kategori_keluar(text):
    low = text.lower()
    if low.startswith("refund"):
        return "Refund"
    if "pantri" in low or "filter air" in low:
        return "Duit Pantri"
    if "lawatan" in low or "trip" in low or "deposit" in low:
        return "Lawatan / Trip"
    if "baju" in low:
        return "Pakaian"
    if "pskpp" in low or "buku perancangan" in low:
        return "Langganan / Bahan"
    if any(k in low for k in ("majlis", "hadiah", "jamuan", "hari guru", "hamper", "boling",
                              "kahwin", "sumbangan", "minuman", "makanan")):
        return "Majlis & Hadiah"
    return "Pelbagai"


def month_tokens(butiran, rate):
    """[(bulan, amaun_separa|None)] daripada kurungan bulan; [] jika tiada."""
    for grp in re.findall(r"\(([^)]*)\)", butiran):
        rest = re.sub(r"kelab|lawatan|yuran|setengah\s+bulan|[\d\s,&\[\]]", "", grp, flags=re.I)
        if rest or not re.search(r"\d", grp):
            continue  # bukan kurungan bulan (cth "Lamb Chop", "Hutang Lawatan")
        out, g = [], grp
        half = re.search(r"setengah\s+bulan\s+(\d+)", g, re.I)
        if half:
            out.append((int(half.group(1)), rate / 2))
            g = g[:half.start()] + g[half.end():]
        g = re.sub(r"kelab|lawatan|yuran|&", " ", g, flags=re.I)
        out = [(int(n), float(a) if a else None) for n, a in re.findall(r"(\d+)(?:\[(\d+)\])?", g)] + out
        return out
    return []


def next_free(claimed, start=START):
    y, m = start
    while (y, m) in claimed:
        y, m = (y + 1, 1) if m == 12 else (y, m + 1)
    return y, m


MULA = {"Aiman": (2025, 12)}  # Aiman masuk bulan 12 (tab "Conclusion")


def year_options(ahli, tarikh, m):
    """Tahun yang munasabah bagi satu bulan: tidak sebelum Mula ahli, antara -8 dan +6 bulan
    daripada tarikh bayaran (longgarkan ke -15..+9 jika tiada), dan tidak melepasi 2026."""
    lo = MULA.get(ahli, START)
    pay = tarikh.year * 12 + tarikh.month - 1
    for past, ahead in ((8, 6), (15, 9)):
        ys = [y for y in (2025, 2026) if (y, m) >= lo and -past <= y * 12 + m - 1 - pay <= ahead]
        if ys:
            return ys
    return []


def parse(ws):
    rows, cur, in_yuran = [], None, False
    for r in range(2, ws.max_row + 1):
        a, b, c, d, e = (ws.cell(r, i).value for i in range(1, 6))
        if a is not None:
            cur, in_yuran = to_date(a), False
        if b is None:
            continue
        if b == "Yuran" and c == "Kelab":
            in_yuran = True
            continue
        if b == "Baki daripada 2025":
            continue  # 2025 dimasukkan penuh, jadi baki bawa-ke-hadapan tidak perlu
        b = str(b).strip()
        kelab = float(c) if isinstance(c, (int, float)) else 0.0
        lawatan = float(d) if isinstance(d, (int, float)) else 0.0
        keluar = float(e) if isinstance(e, (int, float)) else 0.0
        ahli = find_ahli(b)
        base = dict(tarikh=cur, ahli=ahli, butiran=b, bulan="", kelab=0.0, lawatan=0.0, lain=0.0, keluar=0.0)
        if keluar:
            rows.append({**base, "kategori": kategori_keluar(b), "keluar": keluar})
        if in_yuran and ahli and not b.lower().startswith(("majlis", "sumbangan")):
            if kelab or lawatan:
                rows.append({**base, "kategori": "Yuran", "kelab": kelab, "lawatan": lawatan, "_raw": True})
        elif kelab or lawatan:
            rows.append({**base, "kategori": kategori_masuk(b), "lain": kelab + lawatan})
    return rows


def fill_months(rows, tab=None):
    """Isi lajur Bulan. tab = {(ahli, 'kelab'|'lawatan', 'YYYY-MM'): amaun} daripada tab Yuran Kelab
    (bukti tambahan bila tahun tidak jelas). Mengembalikan (inferred, ambiguous)."""
    tab = tab or {}
    claimed = {(a, k): set() for a in AHLI for k in ("kelab", "lawatan")}
    inferred, ambiguous = [], []
    for x in sorted((x for x in rows if x.get("_raw")), key=lambda x: x["tarikh"]):
        cats = [k for k in ("kelab", "lawatan") if x[k]]
        rate = RATE[cats[0]]
        toks = month_tokens(x["butiran"], rate)
        out, status = [], "pasti"
        if toks:
            pay = x["tarikh"].year * 12 + x["tarikh"].month - 1
            for m, amt in toks:
                ys = year_options(x["ahli"], x["tarikh"], m)
                if len(ys) == 1:
                    y = ys[0]
                elif not ys:
                    y, status = (2025 if m >= 6 else 2026), "tiada-calon"
                else:
                    hit = [yy for yy in ys if any(tab.get((x["ahli"], k, f"{yy}-{m:02d}")) for k in cats)]
                    if len(hit) == 1:
                        y = hit[0]  # tab anda sudah tanda bayar pada tahun ini
                    else:
                        y = min(ys, key=lambda yy: abs(yy * 12 + m - 1 - pay))
                        ambiguous.append((x["tarikh"].isoformat(), x["butiran"], cats, f"bulan {m}", [f"{yy}-{m:02d}" for yy in ys]))
                out.append((y, m, amt))
        else:  # tiada bulan ditulis -> JANGAN teka; ditanda untuk disemak pengguna
            inferred.append((x["tarikh"].isoformat(), x["butiran"], cats, x[cats[0]]))
            x["_toks"], x["bulan"], x["_status"] = [], "", "tanpa-bulan"
            continue
        for k in cats:
            claimed[(x["ahli"], k)].update((y, m) for y, m, _ in out)
        x["bulan"] = ", ".join(f"{y}-{m:02d}" + (f":{a:g}" if a else "") for y, m, a in out)
        x["_toks"], x["_status"] = list(out), status
    return inferred, ambiguous


def read_tab(wb):
    """Tab 'Yuran Kelab' lama -> {(ahli, 'kelab'|'lawatan', 'YYYY-MM'): amaun}. Digunakan sebagai bukti tahun."""
    ws, tab, y, prev, cur = wb["Yuran Kelab"], {}, 2025, 0, None
    for r in range(1, ws.max_row + 1):
        v = ws.cell(r, 1).value
        if isinstance(v, str) and v.startswith("Bulan "):
            m = int(v.split()[1])
            y += 1 if m < prev else 0
            prev, cur = m, f"{y}-{m:02d}"
        elif v and cur:
            for k, col in (("kelab", 2), ("lawatan", 3)):
                a = ws.cell(r, col).value
                if isinstance(a, (int, float)) and a > 0:
                    tab[(str(v), k, cur)] = float(a)
    return tab


def mi(y, m):
    return y * 12 + m - 1


def runs_of(tokens):
    """[(y,m,amt)] -> senarai (mula(y,m), akhir(y,m), n, amaun_separa|None). Bulan berturutan digabung;
    bulan separa sentiasa sebaris sendiri."""
    out, cur = [], None
    for y, m, amt in sorted(tokens, key=lambda t: mi(t[0], t[1])):
        if amt is None and cur and cur["amt"] is None and mi(y, m) == mi(*cur["end"]) + 1:
            cur["end"], cur["n"] = (y, m), cur["n"] + 1
        else:
            cur = {"start": (y, m), "end": (y, m), "n": 1, "amt": amt}
            out.append(cur)
    return out


MON = ["Jan", "Feb", "Mac", "Apr", "Mei", "Jun", "Jul", "Ogo", "Sep", "Okt", "Nov", "Dis"]


def label(r):
    (y1, m1), (y2, m2) = r["start"], r["end"]
    if (y1, m1) == (y2, m2):
        t = f"{MON[m1 - 1]} {y1 % 100}"
    elif y1 == y2:
        t = f"{MON[m1 - 1]}\u2013{MON[m2 - 1]} {y1 % 100}"
    else:
        t = f"{MON[m1 - 1]} {y1 % 100}\u2013{MON[m2 - 1]} {y2 % 100}"
    return t + (" separa" if r["amt"] is not None else "")


def extra_note(butiran):
    """Nota dalam kurungan yang bukan bulan, cth '(Hutang Lawatan)'."""
    notes = [g for g in re.findall(r"\(([^)]*)\)", butiran) if re.search(r"hutang|lamb|boling", g, re.I)]
    return (" - " + ", ".join(notes)) if notes else ""


def split_rows(rows):
    """Satu baris yuran dengan bulan -> satu baris setiap larian bulan. Jumlah kekal tepat."""
    out, notes = [], []
    for x in rows:
        if not x.get("_toks"):
            x["mula"] = x["akhir"] = ""
            out.append(x)
            continue
        runs = runs_of(x["_toks"])
        cats = [k for k in ("kelab", "lawatan") if x[k]]
        nominal = {k: [(r["amt"] if r["amt"] is not None else RATE[k]) * r["n"] for r in runs] for k in cats}
        mismatch = any(abs(sum(nominal[k]) - x[k]) > 0.01 for k in cats)
        why = "; ".join(f"{k} RM{x[k]:g} bukan RM{sum(nominal[k]):g}" for k in cats if abs(sum(nominal[k]) - x[k]) > 0.01)
        if len(runs) == 1:
            r = runs[0]
            out.append({**x, "mula": "%d-%02d" % r["start"], "akhir": "%d-%02d" % r["end"],
                        "butiran": f"{x['ahli']} ({label(r)}){extra_note(x['butiran'])}",
                        "_flag": f"Semak: {why}" if mismatch else ""})
            if mismatch:
                notes.append((x["tarikh"].isoformat(), x["butiran"], "jumlah tidak sepadan kadar x bulan"))
            continue
        for i, r in enumerate(runs):
            row = {**x, "mula": "%d-%02d" % r["start"], "akhir": "%d-%02d" % r["end"], "_flag": "",
                   "butiran": f"{x['ahli']} ({label(r)}){extra_note(x['butiran'])}"}
            for k in cats:
                tot = sum(nominal[k]) or 1
                row[k] = round(x[k] * nominal[k][i] / tot, 2)
            out.append(row)
        for k in cats:  # betulkan sen akibat pembundaran pada baris terakhir
            diff = round(x[k] - sum(o[k] for o in out[-len(runs):]), 2)
            out[-1][k] = round(out[-1][k] + diff, 2)
        if mismatch:
            for o in out[-len(runs):]:
                o["_flag"] = f"Semak: {why}"
            notes.append((x["tarikh"].isoformat(), x["butiran"], "jumlah tidak sepadan kadar x bulan (dibahagi ikut nisbah)"))
    return out, notes


def main(src, dst):
    wb = openpyxl.load_workbook(src, data_only=True)
    rows = parse(wb["2025"]) + parse(wb["2026"])
    rows.sort(key=lambda x: x["tarikh"])  # sort stabil: susunan asal dalam hari yang sama kekal
    inferred, ambiguous = fill_months(rows, read_tab(wb))
    rows, notes = split_rows(rows)
    out = []
    for x in rows:
        flag = x.get("_flag", "")
        if x.get("_status") == "tanpa-bulan":
            flag = "Bulan?"
        if x["ahli"] == "Hasnol" and x["tarikh"].isoformat() == "2025-09-26" and x["lawatan"] == 80:
            flag = "Kes khas: bayar RM80 (nota tab lama)"
        out.append(dict(tarikh=x["tarikh"].isoformat(), butiran=x["butiran"], ahli=x["ahli"], kategori=x["kategori"],
                        mula=x.get("mula", ""), akhir=x.get("akhir", ""), kelab=x["kelab"], lawatan=x["lawatan"],
                        lain=x["lain"], keluar=x["keluar"], nota=flag))
    masuk = sum(x["kelab"] + x["lawatan"] + x["lain"] for x in out)
    keluar = sum(x["keluar"] for x in out)
    print(f"{len(out)} baris | masuk {masuk:.2f} | keluar {keluar:.2f} | baki {masuk - keluar:.2f}")
    print(f"{len(ambiguous)} bulan tahun tidak pasti (diselesaikan guna tab Yuran Kelab atau bulan terdekat):")
    for i in ambiguous:
        print("  ", *i)
    print(f"{len(inferred)} baris yuran TANPA bulan (dibiarkan kosong, ditanda 'Bulan?'):")
    for i in inferred:
        print("  ", *i)
    print(f"{len(notes)} baris jumlah tidak sepadan:")
    for i in notes:
        print("  ", *i)
    json.dump(out, open(dst, "w"), ensure_ascii=False, indent=1)


if __name__ == "__main__":
    main(*sys.argv[1:3])
