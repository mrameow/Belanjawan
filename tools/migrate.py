"""Tukar helaian lama (2025 + 2026) kepada senarai transaksi yang bersih.

Guna:  python3 -I tools/migrate.py <Belanjawan_Kelab_Guru.xlsx> <output.json>

Format lama: satu blok tarikh -> beberapa baris "Nama (bulan)" dengan lajur
Kelab / Lawatan, dan baris perbelanjaan di lajur Out. Skrip ini memecahkannya
kepada satu baris = satu transaksi (Tarikh, Jenis, Kategori, Ahli, Butiran, Jumlah).
"""
import datetime as dt
import json
import re
import sys

import openpyxl

AHLI = ["Amir", "Asnawi", "Bang Mie", "Boss Khairil", "Boss Wan", "Chang", "Hadi",
        "Hasnol", "Joe", "Khairiri", "Kuchai", "Suhaimi", "Ustaz", "Wafi", "Aiman"]
ALIAS = {"Boss": "Boss Wan"}  # helaian 2025 tulis "Boss (6)"


def to_date(v, year):
    """Excel menukar 'dd/mm' kepada tarikh mm/dd bila dd <= 12 -> tukar balik."""
    if isinstance(v, dt.datetime):
        return dt.date(v.year, v.day, v.month)  # swap bulan/hari
    if isinstance(v, str):
        d, m, y = re.split(r"/", v.strip())
        return dt.date(int(y), int(m), int(d))
    raise ValueError(v)


def find_ahli(text):
    t = text.strip()
    m = re.match(r"^([A-Za-z ]+?)\s*(\(|$)", t)
    head = m.group(1).strip() if m else t
    head = ALIAS.get(head, head)
    if head in AHLI:
        return head
    # "Majlis KKGS - Hasnol (Lamb Chop)" / "Bayaran ... - Asnawi"
    m = re.search(r"-\s*([A-Za-z ]+?)\s*(\(|$)", t)
    if m:
        tail = ALIAS.get(m.group(1).strip(), m.group(1).strip())
        if tail in AHLI:
            return tail
    return ""


def kategori_masuk(text, in_yuran_block, col):
    low = text.lower()
    if low.startswith("majlis"):
        return "Majlis"
    if "sumbangan" in low:
        return "Sumbangan"
    if in_yuran_block:
        return "Yuran Kelab" if col == "C" else "Yuran Lawatan"
    return "Lain-lain"


def kategori_keluar(text):
    low = text.lower()
    if "pantri" in low or "filter air" in low:
        return "Duit Pantri"
    if low.startswith("refund"):
        return "Refund"
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


def parse(ws, year):
    out, cur = [], None
    in_yuran = False
    for r in range(2, ws.max_row + 1):
        a, b, c, d, e = (ws.cell(r, i).value for i in range(1, 6))
        if a is not None:
            cur = to_date(a, year)
            in_yuran = False
        if b is None:
            continue
        if b == "Yuran" and c == "Kelab":
            in_yuran = True
            continue
        if b == "Baki daripada 2025":
            continue  # 2025 dimasukkan penuh, jadi baki bawa-ke-hadapan tidak perlu
        b = str(b).strip()
        if e is not None:
            out.append(dict(tarikh=cur, jenis="Keluar", kategori=kategori_keluar(b),
                            ahli=find_ahli(b), butiran=b, jumlah=float(e)))
        for col, val in (("C", c), ("D", d)):
            if isinstance(val, (int, float)):
                kat = kategori_masuk(b, in_yuran, col)
                if kat.startswith("Yuran") or col == "C":
                    out.append(dict(tarikh=cur, jenis="Masuk", kategori=kat, ahli=find_ahli(b),
                                    butiran=b, jumlah=float(val)))
                else:
                    raise ValueError((r, b, val))
    return out


def main(src, dst):
    wb = openpyxl.load_workbook(src, data_only=True)
    rows = parse(wb["2025"], 2025) + parse(wb["2026"], 2026)
    rows.sort(key=lambda x: x["tarikh"])  # sort stabil, kekal susunan asal dalam hari yang sama
    for x in rows:
        x["tarikh"] = x["tarikh"].isoformat()
    masuk = sum(x["jumlah"] for x in rows if x["jenis"] == "Masuk")
    keluar = sum(x["jumlah"] for x in rows if x["jenis"] == "Keluar")
    print(f"{len(rows)} transaksi | masuk {masuk:.2f} | keluar {keluar:.2f} | baki {masuk - keluar:.2f}")
    json.dump(rows, open(dst, "w"), ensure_ascii=False, indent=1)


if __name__ == "__main__":
    main(*sys.argv[1:3])
