#!/usr/bin/env python3
"""ALU işlem günlüğü: veri_alu.xlsx + cevap_anahtari.xlsx üretici.

Tüm ALU hesapları yalnızca NOT ve AND kapılarıyla yapılır (ALU sınıfı).
Kullanım:  python alu_uret.py [çıktı_klasörü]      (varsayılan: betiğin klasörü)
Gereken:   openpyxl  (+ doğrulama için pandas, numpy)
"""
import bisect
import datetime as dt
import math
import random
import sys
from collections import Counter, defaultdict
from pathlib import Path

from openpyxl import Workbook
from openpyxl.formatting.rule import CellIsRule
from openpyxl.styles import Alignment, Border, Font, PatternFill, Side
from openpyxl.utils import get_column_letter
from openpyxl.worksheet.properties import PageSetupProperties

# ───────────────────────────── sabitler ─────────────────────────────
SEED0 = 2026
N = 5000
W = 8
BASLA = dt.datetime(2026, 8, 1)
GUN = 92                                   # 1 Ağu – 31 Eki
TUM_SN = GUN * 86400
SESSIZ = 3                                 # ~30 saat hiç işlem yapmayan çekirdek
OLAY_CEKIRDEKLERI = [1, 2, 2, 4, 4, 5, 5, 6, 7, 7, 8, 8]   # 12 termal olay (3 numara yok)
HATA_ORANI = 0.40
HATA_BIT_AGIRLIK = [10, 10, 10, 10, 10, 30, 10, 10]        # bit 5 ağırlıklı (0 = en sağ)
TURLER = ["TOPLA", "ÇIKAR", "AND", "OR", "XOR", "NOT"]
AGIRLIK = [30, 20, 15, 15, 15, 5]
KAPI_BEKLENEN = {"NOT": (8, 0), "AND": (0, 8), "OR": (24, 8),
                 "XOR": (32, 24), "TOPLA": (88, 72), "ÇIKAR": (96, 72)}
CEKIRDEKLER = [(1, "AS-P1 Performans", 4.2), (2, "AS-P1 Performans", 4.0),
               (3, "AS-P2 Performans", 3.8), (4, "AS-P2 Performans", 3.6),
               (5, "AS-E1 Verimlilik", 3.0), (6, "AS-E1 Verimlilik", 2.8),
               (7, "AS-E2 Verimlilik", 2.6), (8, "AS-E2 Verimlilik", 2.4)]
HIZ = {c: g for c, _, g in CEKIRDEKLER}
# başlangıç tarihi, nJ/kapı (yüzde birlik tamsayı → toplamlar tam 2 ondalık çıkar)
TARIFE = [(dt.date(2026, 8, 1), 28, "Dönem 1"), (dt.date(2026, 8, 24), 31, "Dönem 2"),
          (dt.date(2026, 9, 15), 26, "Dönem 3"), (dt.date(2026, 10, 12), 35, "Dönem 4")]
TARIFE_BAS = [t[0] for t in TARIFE]

BASLIKLAR = ["İşlem No", "Tarih-Saat", "Çekirdek No", "İşlem Türü", "A (ondalık)", "B (ondalık)",
             "A (ikilik)", "B (ikilik)", "Sonuç (ikilik)", "Taşma", "Taşıma Zinciri",
             "Kapı Sayısı", "Süre (ns)"]


def zaman(s):
    return BASLA + dt.timedelta(seconds=s)


# ───────────────────── yalnızca NOT ve AND'den ALU ─────────────────────
class ALU:
    def __init__(self):
        self.n_not = self.n_and = 0

    def NOT(self, a):
        self.n_not += 1
        return 1 - a

    def AND(self, a, b):
        self.n_and += 1
        return a & b

    def OR(self, a, b):                       # 3 NOT + 1 AND
        return self.NOT(self.AND(self.NOT(a), self.NOT(b)))

    def XOR(self, a, b):                      # 4 NOT + 3 AND
        return self.AND(self.OR(a, b), self.NOT(self.AND(a, b)))

    def tam_toplayici(self, a, b, cin):       # 11 NOT + 9 AND = 20 kapı
        x = self.XOR(a, b)
        return self.XOR(x, cin), self.OR(self.AND(a, b), self.AND(x, cin))

    def topla(self, a, b, cin):
        carry, toplam, cikislar = cin, [], []
        for i in range(W):
            t, carry = self.tam_toplayici(a[i], b[i], carry)
            toplam.append(t)
            cikislar.append(carry)
        return toplam, cikislar


def bitler(x):
    return [(x >> i) & 1 for i in range(W)]


def tam(bl):
    return sum(b << i for i, b in enumerate(bl))


def ikilik(x):
    return format(x, "08b")


def en_uzun_seri(c):
    en, cur = 0, 0
    for v in c:
        cur = cur + 1 if v else 0
        en = max(en, cur)
    return en


def calistir(tur, A, B):
    alu = ALU()
    a = bitler(A)
    b = bitler(B) if B is not None else None
    tasma, zincir = False, 0
    if tur == "NOT":
        r = [alu.NOT(x) for x in a]
    elif tur == "AND":
        r = [alu.AND(x, y) for x, y in zip(a, b)]
    elif tur == "OR":
        r = [alu.OR(x, y) for x, y in zip(a, b)]
    elif tur == "XOR":
        r = [alu.XOR(x, y) for x, y in zip(a, b)]
    elif tur == "TOPLA":
        r, c = alu.topla(a, b, 0)
        tasma, zincir = c[-1] == 1, en_uzun_seri(c)
    else:  # ÇIKAR: A + NOT(B) + 1  (taşıma çıkışı yoksa ödünç var → A<B; tel okuması, kapı değil)
        r, c = alu.topla(a, [alu.NOT(y) for y in b], 1)
        tasma = c[-1] == 0
    return dict(sonuc=tam(r), tasma=tasma, zincir=zincir, n_not=alu.n_not, n_and=alu.n_and)


def beklenen(tur, A, B):
    if tur == "TOPLA":
        return (A + B) & 255
    if tur == "ÇIKAR":
        return (A - B) & 255
    if tur == "AND":
        return A & B
    if tur == "OR":
        return A | B
    if tur == "XOR":
        return A ^ B
    return 255 - A


def kontrol_et(kosul, mesaj):
    if not kosul:
        raise RuntimeError("MOTOR HATASI: " + mesaj)


def motor_testi(rng):
    """1-bit tablolar tam, 8-bit işlemler rastgele + köşe değerlerle Python işlemleriyle kıyaslanır."""
    g = ALU()
    for a in (0, 1):
        kontrol_et(g.NOT(a) == 1 - a, "NOT")
        for b in (0, 1):
            kontrol_et(g.AND(a, b) == (a & b) and g.OR(a, b) == (a | b) and g.XOR(a, b) == (a ^ b), "1-bit")
            for c in (0, 1):
                s, co = g.tam_toplayici(a, b, c)
                kontrol_et(s == (a + b + c) % 2 and co == (a + b + c) // 2, "tam toplayıcı")
    cift = [(0, 0), (255, 255), (255, 1), (1, 255), (0, 255), (128, 128), (127, 1)]
    cift += [(rng.randrange(256), rng.randrange(256)) for _ in range(2000)]
    for tur in TURLER:
        for A, B in cift:
            B2 = None if tur == "NOT" else B
            m = calistir(tur, A, B2)
            kontrol_et(m["sonuc"] == beklenen(tur, A, B2), f"{tur} {A},{B}")
            kontrol_et((m["n_not"], m["n_and"]) == KAPI_BEKLENEN[tur], f"kapı sayısı {tur}")


# ───────────────────────────── veri üretimi ─────────────────────────────
def uret(seed):
    rng = random.Random(seed)
    motor_testi(rng)

    s_bas = rng.randrange(20 * 86400, 70 * 86400, 60)
    s_bit = s_bas + 30 * 3600

    olaylar = []
    for c in OLAY_CEKIRDEKLERI:
        while True:
            bas = rng.randrange(86400, 90 * 86400, 60)
            sure = rng.randint(30, 90)
            bit = bas + sure * 60
            if all(bit + 21600 < o["bas"] or o["bit"] + 21600 < bas for o in olaylar if o["c"] == c):
                break
        olaylar.append(dict(c=c, bas=bas, bit=bit, sure=sure))
    olaylar.sort(key=lambda o: (o["bas"], o["c"]))
    olay_c = defaultdict(list)
    for o in olaylar:
        olay_c[o["c"]].append(o)

    def pencerede(c, s):
        return any(o["bas"] <= s <= o["bit"] for o in olay_c[c])

    def sinira_yakin(c, s):                      # Excel metin-yuvarlama riskine karşı ±3 sn tampon
        return any(abs(s - o["bas"]) <= 3 or abs(s - o["bit"]) <= 3 for o in olay_c[c])

    kullanilan, ops = set(), []
    for o in olaylar:                            # termal yük: pencere başına ~12 işlem
        for _ in range(rng.randint(9, 15)):
            while True:
                s = rng.randint(o["bas"] + 4, o["bit"] - 4)
                if s not in kullanilan:
                    break
            kullanilan.add(s)
            ops.append((s, o["c"]))
    while len(ops) < N:                          # arka plan işlemleri
        s, c = rng.randrange(TUM_SN), rng.randint(1, 8)
        if s in kullanilan or sinira_yakin(c, s) or (c == SESSIZ and s_bas <= s < s_bit):
            continue
        kullanilan.add(s)
        ops.append((s, c))
    ops.sort()

    rows = []
    for no, (s, c) in enumerate(ops, 1):
        tur = rng.choices(TURLER, AGIRLIK)[0]
        A = rng.randrange(256)
        B = None if tur == "NOT" else rng.randrange(256)
        m = calistir(tur, A, B)
        bek = beklenen(tur, A, B)
        kontrol_et(m["sonuc"] == bek, f"günlük satırı {no}: {tur} {A} {B}")
        kontrol_et((m["n_not"], m["n_and"]) == KAPI_BEKLENEN[tur], f"kapı sayısı satır {no}")
        if tur == "TOPLA":
            kontrol_et(m["tasma"] == (A + B > 255), f"taşma satır {no}")
        if tur == "ÇIKAR":
            kontrol_et(m["tasma"] == (A < B), f"ödünç satır {no}")
        kapi = m["n_not"] + m["n_and"]
        pen = pencerede(c, s)
        temel = (kapi * 0.04 + m["zincir"] * 0.9) / HIZ[c]
        if pen:
            temel *= 1.35
        sure = round(max(temel + rng.gauss(0, 0.10), 0.05), 2)
        son, poz = bek, None
        if pen and rng.random() < HATA_ORANI:
            poz = rng.choices(range(W), HATA_BIT_AGIRLIK)[0]
            son = bek ^ (1 << poz)
        rows.append(dict(no=f"ISL-{no:05d}", s=s, c=c, tur=tur, A=A, B=B, ab=ikilik(A),
                         bb=ikilik(B) if B is not None else None, sonuc=ikilik(son),
                         tasma="EVET" if m["tasma"] else "HAYIR", zincir=m["zincir"],
                         kapi=kapi, sure=sure, pen=pen, poz=poz, bek=bek))
    return rows, olaylar, (s_bas, s_bit)


# ───────────────────────────── cevaplar ─────────────────────────────
def birim(s):
    return TARIFE[bisect.bisect_right(TARIFE_BAS, zaman(s).date()) - 1][1]


def cevaplari_hesapla(rows):
    a = {}
    a["q1"] = sum(r["sonuc"].count("1") for r in rows)
    a["q2"] = round(sum(r["kapi"] * birim(r["s"]) for r in rows if zaman(r["s"]).month == 10) / 100, 2)

    def isaretli(x):
        return x - 256 if x > 127 else x
    a["q3"] = sum(1 for r in rows if r["tur"] == "TOPLA"
                  and not -128 <= isaretli(r["A"]) + isaretli(r["B"]) <= 127)
    a["q4"] = sum(r["pen"] for r in rows)
    ts = [r["s"] for r in rows]
    a["q5"] = max(bisect.bisect_left(ts, t + 3600) - i for i, t in enumerate(ts))
    hatali = [r for r in rows if int(r["sonuc"], 2) != r["bek"]]
    a["q6"] = len(hatali)
    a["q7"] = round(sum(1 for r in hatali if r["pen"]) / a["q4"] * 100, 2)
    a["bit_dagilim"] = Counter(r["poz"] for r in hatali)
    a["q8"] = a["bit_dagilim"].most_common(1)[0][0]
    gaps = []
    for c in range(1, 9):
        t = [r["s"] for r in rows if r["c"] == c]
        gaps += [(y - x, c) for x, y in zip(t, t[1:])]
    gaps.sort(reverse=True)
    a["gaps"] = gaps[:3]
    a["q9"] = round(gaps[0][0] / 3600, 2)
    pts = [(r["zincir"], r["sure"]) for r in rows if r["c"] == 5 and r["tur"] == "TOPLA" and not r["pen"]]
    n = len(pts)
    mx, my = math.fsum(p[0] for p in pts) / n, math.fsum(p[1] for p in pts) / n
    sxy = math.fsum((x - mx) * (y - my) for x, y in pts)
    sxx = math.fsum((x - mx) ** 2 for x, _ in pts)
    syy = math.fsum((y - my) ** 2 for _, y in pts)
    a["q10_ham"] = sxy / sxx
    a["q10"] = round(a["q10_ham"], 2)
    a["q10_n"], a["q10_r2"] = n, sxy ** 2 / (sxx * syy)
    a["hatali"] = hatali
    return a


def yuvarlama_guvenli(x, nd=2):
    y = x * 10 ** nd
    return abs((y - math.floor(y)) - 0.5) > 1e-3


def kontrol(rows, a):
    if a["q4"] < 100 or a["q6"] < 30:
        return "pencere içi işlem / hata sayısı az"
    mc = a["bit_dagilim"].most_common(2)
    if mc[0][1] - mc[1][1] < 3:
        return "Soru 8'de eşitlik riski"
    (g1, c1), (g2, _) = a["gaps"][0], a["gaps"][1]
    if c1 != SESSIZ or g1 - g2 < 3600:
        return "Soru 9'da en uzun bekleme belirsiz"
    sureler = sorted((r["sure"] for r in rows), reverse=True)
    if sureler[0] == sureler[1]:
        return "en uzun süre eşit"
    if a["q10_n"] < 100:
        return "Soru 10 için az veri"
    if abs(a["q2"] - round(a["q2"])) < 1e-9:
        return "Soru 2 ondalıksız çıktı"
    if not all(yuvarlama_guvenli(a[k]) for k in ("q7",)) or not yuvarlama_guvenli(a["q9"] * 0 + a["gaps"][0][0] / 3600) \
            or not yuvarlama_guvenli(a["q10_ham"]):
        return "yuvarlama sınırında cevap"
    return None


# ───────────────────────── Excel biçim yardımcıları ─────────────────────────
F = "Arial"
F_NORM = Font(name=F, size=10)
F_KALIN = Font(name=F, size=10, bold=True)
F_MONO = Font(name="Consolas", size=10)
F_BASLIK = Font(name=F, size=10, bold=True, color="FFFFFF")
F_BASLIK_B = Font(name=F, size=14, bold=True, color="1F3864")
F_NOT = Font(name=F, size=9, italic=True, color="595959")
D_BASLIK = PatternFill("solid", fgColor="1F3864")
D_ZEBRA = PatternFill("solid", fgColor="F3F6FB")
D_ACIK = PatternFill("solid", fgColor="DDE6F3")
KENAR = Border(*[Side(style="thin", color="C9D1DE")] * 4)
FMT_DT = "dd.mm.yyyy hh:mm:ss"


def stil_baslik(ws, satir, n, yukseklik=32):
    al = Alignment(horizontal="center", vertical="center", wrap_text=True)
    for j in range(1, n + 1):
        c = ws.cell(satir, j)
        c.font, c.fill, c.alignment, c.border = F_BASLIK, D_BASLIK, al, KENAR
    ws.row_dimensions[satir].height = yukseklik


def tablo_yaz(ws, basliklar, satirlar, genislik, bicim=None, hiza=None, mono=(), duz_metin=(),
              zebra=True, filtre=True, dondur="A2", sekme=None):
    bicim, hiza = bicim or {}, hiza or {}
    for j, b in enumerate(basliklar, 1):
        ws.cell(1, j, b)
    stil_baslik(ws, 1, len(basliklar))
    al = {j: Alignment(horizontal=hiza.get(j, "center"), vertical="center", wrap_text=j in duz_metin)
          for j in range(1, len(basliklar) + 1)}
    son = 1
    for i, satir in enumerate(satirlar, 2):
        son = i
        for j, v in enumerate(satir, 1):
            c = ws.cell(i, j, v)
            if j in duz_metin and isinstance(v, str):
                c.data_type = "s"                      # "=..." ile başlayan metin formül sayılmasın
                try:
                    c.quotePrefix = True
                except AttributeError:
                    pass
            c.font = F_MONO if j in mono else F_NORM
            c.border, c.alignment = KENAR, al[j]
            if j in bicim:
                c.number_format = bicim[j]
            if zebra and i % 2 == 1:
                c.fill = D_ZEBRA
    for j, g in enumerate(genislik, 1):
        ws.column_dimensions[get_column_letter(j)].width = g
    if dondur:
        ws.freeze_panes = dondur
    if filtre and son > 1:
        ws.auto_filter.ref = f"A1:{get_column_letter(len(basliklar))}{son}"
    if sekme:
        ws.sheet_properties.tabColor = sekme
    ws.page_setup.orientation = "landscape"
    ws.page_setup.fitToWidth, ws.page_setup.fitToHeight = 1, 0
    ws.sheet_properties.pageSetUpPr = PageSetupProperties(fitToPage=True)
    ws.print_title_rows = "1:1"
    return son


# ───────────────────────────── veri_alu.xlsx ─────────────────────────────
def yaz_islemler(ws, rows):
    satirlar = [[r["no"], zaman(r["s"]), r["c"], r["tur"], r["A"], r["B"], r["ab"], r["bb"], r["sonuc"],
                 r["tasma"], r["zincir"], r["kapi"], r["sure"]] for r in rows]
    son = tablo_yaz(ws, BASLIKLAR, satirlar,
                    [12, 21, 12, 12, 12, 12, 13, 13, 15, 9, 15, 12, 11],
                    bicim={2: FMT_DT, 7: "@", 8: "@", 9: "@", 13: "0.00"},
                    hiza={1: "left"}, mono=(7, 8, 9), dondur="B2", sekme="1F3864")
    renkler = {"TOPLA": "DDEBF7", "ÇIKAR": "FCE4D6", "AND": "E2EFDA",
               "OR": "FFF2CC", "XOR": "EADCF4", "NOT": "EDEDED"}
    for tur, renk in renkler.items():
        ws.conditional_formatting.add(f"D2:D{son}", CellIsRule(
            operator="equal", formula=[f'"{tur}"'],
            fill=PatternFill(start_color=renk, end_color=renk, fill_type="solid")))
    ws.conditional_formatting.add(f"J2:J{son}", CellIsRule(
        operator="equal", formula=['"EVET"'], font=Font(name=F, bold=True, color="C00000")))


def yaz_dogruluk(ws):
    g = ALU()
    bloklar = [
        ("NOT  (temel kapı)", ["A", "NOT A"],
         [(a, g.NOT(a)) for a in (0, 1)], "2E75B6"),
        ("AND  (temel kapı)", ["A", "B", "A AND B"],
         [(a, b, g.AND(a, b)) for a in (0, 1) for b in (0, 1)], "2E75B6"),
        ("OR  =  NOT( AND( NOT A , NOT B ) )", ["A", "B", "NOT A", "NOT B", "AND(NOT A, NOT B)", "A OR B"],
         [(a, b, g.NOT(a), g.NOT(b), g.AND(g.NOT(a), g.NOT(b)), g.OR(a, b)) for a in (0, 1) for b in (0, 1)],
         "548235"),
        ("XOR  =  AND( OR(A,B) , NOT( AND(A,B) ) )", ["A", "B", "A OR B", "NOT(A AND B)", "A XOR B"],
         [(a, b, g.OR(a, b), g.NOT(g.AND(a, b)), g.XOR(a, b)) for a in (0, 1) for b in (0, 1)], "BF8F00"),
        ("Tam toplayıcı  (bit başına 11 NOT + 9 AND = 20 kapı)", ["A", "B", "Cin", "A XOR B", "Toplam", "Cout"],
         [(a, b, c, g.XOR(a, b), *g.tam_toplayici(a, b, c)) for a in (0, 1) for b in (0, 1) for c in (0, 1)],
         "7030A0"),
    ]
    ws.merge_cells("A1:F1")
    ws["A1"] = "Doğruluk tabloları (tümü yalnızca NOT ve AND kapılarından kurulur)"
    ws["A1"].font, ws["A1"].alignment = F_BASLIK_B, Alignment(vertical="center")
    ws.row_dimensions[1].height = 28
    r = 3
    center = Alignment(horizontal="center", vertical="center", wrap_text=True)
    for baslik, kolonlar, satirlar, renk in bloklar:
        ws.merge_cells(start_row=r, start_column=1, end_row=r, end_column=6)
        ws.cell(r, 1, baslik).font = Font(name=F, size=11, bold=True, color="FFFFFF")
        ws.cell(r, 1).alignment = Alignment(vertical="center", indent=1)
        for j in range(1, 7):
            ws.cell(r, j).fill = PatternFill("solid", fgColor=renk)
        ws.row_dimensions[r].height = 22
        r += 1
        for j, k in enumerate(kolonlar, 1):
            c = ws.cell(r, j, k)
            c.font, c.fill, c.alignment, c.border = F_KALIN, D_ACIK, center, KENAR
        ws.row_dimensions[r].height = 30
        r += 1
        for satir in satirlar:
            for j, v in enumerate(satir, 1):
                c = ws.cell(r, j, v)
                c.font, c.alignment, c.border = F_NORM, center, KENAR
                if j == len(satir):
                    c.font, c.fill = F_KALIN, PatternFill("solid", fgColor="FFF2CC")
            r += 1
        r += 1
    for j in range(1, 7):
        ws.column_dimensions[get_column_letter(j)].width = 20
    ws.sheet_view.showGridLines = False
    ws.sheet_properties.tabColor = "7030A0"


def yaz_aciklama(ws, n_olay):
    ws.sheet_view.showGridLines = False
    ws.sheet_properties.tabColor = "7F7F7F"
    ws.merge_cells("A1:D1")
    ws["A1"] = "ALU işlem günlüğü: veri seti açıklaması"
    ws["A1"].font = F_BASLIK_B
    ws.row_dimensions[1].height = 28
    ws.merge_cells("A3:D3")
    ws["A3"] = ("Hikâye: 8 çekirdekli bir işlemcinin 1 Ağustos – 31 Ekim 2026 arasındaki 5.000 işlemlik günlüğü. "
                "Her işlem, yalnızca NOT ve AND kapılarından kurulmuş 8 bitlik bir ALU'da hesaplanmıştır. "
                "Aşırı ısınan çekirdeklerde bazı sonuçlar tek bir bit hatasıyla kaydedilmiştir; "
                "günlükte bu kayıtlar işaretli değildir. Amaç, günlüğü Excel'de doğrulayıp hatanın izini sürmektir.")
    ws["A3"].font = F_NORM
    ws["A3"].alignment = Alignment(wrap_text=True, vertical="top")
    ws.row_dimensions[3].height = 66

    def bolum(r, baslik, kolonlar, satirlar):
        ws.cell(r, 1, baslik).font = Font(name=F, size=12, bold=True, color="1F3864")
        r += 1
        for j, k in enumerate(kolonlar, 1):
            ws.cell(r, j, k)
        stil_baslik(ws, r, 4, 22)
        r += 1
        for satir in satirlar:
            for j, v in enumerate(satir, 1):
                c = ws.cell(r, j, v)
                c.font, c.border = F_NORM, KENAR
                c.alignment = Alignment(wrap_text=True, vertical="center",
                                        horizontal="left" if j in (1, 2, 4) else "center")
                if r % 2 == 0:
                    c.fill = D_ZEBRA
            r += 1
        return r + 1

    r = bolum(5, "Sayfalar", ["Sayfa", "Satır", "", "İçerik"], [
        ("islemler", N, "", "Ana tablo: işlem günlüğü"),
        ("cekirdekler", 8, "", "Çekirdek No, model, üretici, saat hızı (GHz)"),
        ("islem_turleri", 6, "", "İşlem açıklamaları, kapı sayıları, beklenen sonucun Excel karşılığı"),
        ("tarife", len(TARIFE), "", "Kapı başına enerji (nJ); başlangıç tarihine göre aralıklı tablo"),
        ("termal_olaylar", n_olay, "", "Aşırı ısınma pencereleri: Çekirdek No, Başlangıç, Bitiş, Süre (dk)"),
        ("dogruluk_tablosu", "–", "", "NOT, AND ve bunlardan kurulan OR, XOR, tam toplayıcı tabloları"),
        ("aciklama", "–", "", "Bu sayfa")])
    r = bolum(r, "islemler sütunları", ["Sütun", "Ad", "Tür", "Açıklama"], [
        ("A", "İşlem No", "Metin", "ISL-00001 biçiminde; Tarih-Saat sırasına göre"),
        ("B", "Tarih-Saat", "Tarih", "Gerçek Excel tarihi, saniye hassasiyetinde; 1 Ağustos – 31 Ekim 2026"),
        ("C", "Çekirdek No", "Sayı", "1–8"),
        ("D", "İşlem Türü", "Metin", "TOPLA, ÇIKAR, AND, OR, XOR, NOT"),
        ("E", "A (ondalık)", "Sayı", "0–255"),
        ("F", "B (ondalık)", "Sayı", "0–255 (NOT işleminde boş)"),
        ("G", "A (ikilik)", "Metin", "8 karakter; en soldaki bit 7, en sağdaki bit 0"),
        ("H", "B (ikilik)", "Metin", "8 karakter (NOT işleminde boş)"),
        ("I", "Sonuç (ikilik)", "Metin", "Günlüğe kaydedilen sonuç, 8 karakter"),
        ("J", "Taşma", "Metin", "EVET/HAYIR. TOPLA: işaretsiz taşma (A+B>255). ÇIKAR: A<B. Diğerleri: HAYIR"),
        ("K", "Taşıma Zinciri", "Sayı", "TOPLA'da art arda taşıma üreten en uzun bit dizisinin uzunluğu (0–8); diğerlerinde 0"),
        ("L", "Kapı Sayısı", "Sayı", "İşlemde çalışan NOT + AND kapılarının toplamı"),
        ("M", "Süre (ns)", "Sayı", "İşlem süresi; kapı sayısı, taşıma zinciri, çekirdek hızı ve çalışma koşullarına bağlı, ölçüm gürültüsü içerir")])
    ws.column_dimensions["A"].width = 18
    ws.column_dimensions["B"].width = 18
    ws.column_dimensions["C"].width = 10
    ws.column_dimensions["D"].width = 95


def veri_yaz(yol, rows, olaylar):
    wb = Workbook()
    yaz_islemler(wb.active, rows)
    wb.active.title = "islemler"
    tablo_yaz(wb.create_sheet("cekirdekler"), ["Çekirdek No", "Model", "Üretici", "Saat Hızı (GHz)"],
              [(c, m, "Anadolu Silikon", g) for c, m, g in CEKIRDEKLER],
              [14, 22, 20, 16], bicim={4: "0.0"}, filtre=False, sekme="548235")
    tur_satir = [
        ("NOT", "Her bitin tersi (8 bit)", "255-A", "—"),
        ("AND", "Bit bit VE", "BITAND(A;B)", "HAYIR"),
        ("OR", "Bit bit VEYA", "BITOR(A;B)", "HAYIR"),
        ("XOR", "Bit bit özel VEYA", "BITXOR(A;B)", "HAYIR"),
        ("TOPLA", "8 bitlik toplama (tam toplayıcı zinciri)", "MOD(A+B;256)", "A+B > 255"),
        ("ÇIKAR", "A + NOT(B) + 1 ile çıkarma", "MOD(A-B;256)", "A < B")]
    tablo_yaz(wb.create_sheet("islem_turleri"),
              ["İşlem", "Açıklama", "NOT Sayısı", "AND Sayısı", "Toplam Kapı",
               "Beklenen Sonuç (Excel karşılığı)", "Taşma Kuralı"],
              [(t, a, *KAPI_BEKLENEN[t], sum(KAPI_BEKLENEN[t]), x, k) for t, a, x, k in tur_satir],
              [10, 42, 12, 12, 13, 32, 14], hiza={2: "left"}, filtre=False, sekme="BF8F00")
    tablo_yaz(wb.create_sheet("tarife"), ["Başlangıç Tarihi", "Kapı Başına Enerji (nJ)", "Dönem"],
              [(dt.datetime.combine(d, dt.time()), b / 100, ad) for d, b, ad in TARIFE],
              [18, 24, 12], bicim={1: "dd.mm.yyyy", 2: "0.00"}, filtre=False, sekme="C55A11")
    tablo_yaz(wb.create_sheet("termal_olaylar"), ["Çekirdek No", "Başlangıç", "Bitiş", "Süre (dk)"],
              [(o["c"], zaman(o["bas"]), zaman(o["bit"]), o["sure"]) for o in olaylar],
              [14, 20, 20, 12], bicim={2: "dd.mm.yyyy hh:mm", 3: "dd.mm.yyyy hh:mm"},
              filtre=False, sekme="C00000")
    yaz_dogruluk(wb.create_sheet("dogruluk_tablosu"))
    yaz_aciklama(wb.create_sheet("aciklama"), len(olaylar))
    wb.save(yol)


# ───────────────────────────── cevap_anahtari.xlsx ─────────────────────────────
YARDIMCI = {   # N..W  (Excel: İngilizce adlar, virgül ayraç)
    "N": "=IF(E{r}>127,E{r}-256,E{r})",
    "O": "=IF(F{r}>127,F{r}-256,F{r})",
    "P": '=IF(D{r}="TOPLA",IF(OR(N{r}+O{r}<-128,N{r}+O{r}>127),1,0),0)',
    "Q": "=L{r}*VLOOKUP(B{r},tarife!$A$2:$B$5,2,TRUE)",
    "R": '=IF(COUNTIFS(termal_olaylar!$A$2:$A$13,C{r},termal_olaylar!$B$2:$B$13,"<="&B{r},'
         'termal_olaylar!$C$2:$C$13,">="&B{r})>0,1,0)',
    "S": '=COUNTIFS($B$2:$B$5001,">="&(B{r}-0.5/86400),$B$2:$B$5001,"<"&(B{r}+1/24-0.5/86400))',
    "T": '=IF(D{r}="TOPLA",MOD(E{r}+F{r},256),IF(D{r}="ÇIKAR",MOD(E{r}-F{r},256),IF(D{r}="AND",'
         'BITAND(E{r},F{r}),IF(D{r}="OR",BITOR(E{r},F{r}),IF(D{r}="XOR",BITXOR(E{r},F{r}),255-E{r})))))',
    "U": "=IF(BIN2DEC(I{r})<>T{r},1,0)",
    "V": '=IF(U{r}=1,ROUND(LOG(BITXOR(T{r},BIN2DEC(I{r})),2),0),"")',
    "W": '=IF(COUNTIFS($C$2:$C$5001,C{r},$B$2:$B$5001,"<"&(B{r}-0.5/86400))=0,"",(B{r}-MAXIFS($B$2:$B$5001,'
         '$C$2:$C$5001,C{r},$B$2:$B$5001,"<"&(B{r}-0.5/86400)))*24)',
}
YARDIMCI_AD = {"N": "A (işaretli)", "O": "B (işaretli)", "P": "İşaretli taşma", "Q": "Enerji (nJ)",
               "R": "Termal pencerede", "S": "60 dk içindeki işlem", "T": "Beklenen sonuç",
               "U": "Hatalı kayıt", "V": "Ters dönen bit", "W": "Önceki işlemden fark (saat)"}
_KOSUL = '(C2:C5001=5)*(D2:D5001="TOPLA")*(R2:R5001=0)'
CEVAP_FORMUL = [
    (1, "tek hücre", '=SUMPRODUCT(LEN(I2:I5001)-LEN(SUBSTITUTE(I2:I5001,"1","")))', ""),
    (2, "Q sütunu + tek hücre", '=SUMIFS(Q2:Q5001,B2:B5001,">="&DATE(2026,10,1),B2:B5001,"<"&DATE(2026,11,1))',
     "Q: =L2*VLOOKUP(B2,tarife!$A$2:$B$5,2,TRUE)"),
    (3, "N, O, P sütunları + tek hücre", "=SUM(P2:P5001)", ""),
    (4, "R sütunu + tek hücre", "=SUM(R2:R5001)", ""),
    (5, "S sütunu + tek hücre", "=MAX(S2:S5001)",
     "S'de ±0,5 sn tolerans şart: ham B2 / B2+1/24 kıyası Excel'in 15 haneli metin yuvarlaması yüzünden kendini sayabilir ya da atlayabilir"),
    (6, "T, U sütunları + tek hücre", "=SUM(U2:U5001)", ""),
    (7, "tek hücre", "=COUNTIFS(R2:R5001,1,U2:U5001,1)/COUNTIFS(R2:R5001,1)*100", ""),
    (8, "V sütunu + tek hücre", "=MODE(V2:V5001)", "Özet tabloyla da bulunur (Satır = V, Değer = Sayı); LOG sonucu YUVARLA ile tamsayıya çekilir"),
    (9, "W sütunu + tek hücre", "=MAX(W2:W5001)",
     "W'de ±0,5 sn tolerans şart (Soru 5 ile aynı neden); MAXIFS Excel 2019+"),
    (10, "tek hücre (dizi formülü)",
     f'=SLOPE(IF({_KOSUL},M2:M5001),IF({_KOSUL},K2:K5001))', "Eski Excel'de Ctrl+Shift+Enter; alternatif: süz, sonra =SLOPE(M,K)"),
]
SORULAR = [
    (1, "Kolay", "Sonuç (ikilik) sütunundaki toplam 1 biti sayısı", "Tam sayı"),
    (2, "Kolay-orta", "Ekim 2026'daki tüm işlemlerin toplam enerjisi (nJ)", "XXXXX,XX"),
    (3, "Orta", "İşaretli yorumda işaretli taşma oluşan TOPLA işlemi sayısı", "Tam sayı"),
    (4, "Orta", "Kendi çekirdeğinin termal olay penceresi içinde yapılan işlem sayısı", "Tam sayı"),
    (5, "Orta-zor", "Herhangi bir 60 dakikalık dilimdeki en yüksek işlem sayısı", "Tam sayı"),
    (6, "Orta-zor", "Sonucu beklenenden farklı olan kayıt sayısı", "Tam sayı"),
    (7, "Zor", "Termal pencere içindeki işlemlerin yüzde kaçı hatalı", "% XX,XX"),
    (8, "Zor", "En sık bozulan bit konumu (0 = en sağ)", "0–7 tam sayı"),
    (9, "Zor", "Aynı çekirdeğin ardışık iki işlemi arasındaki en uzun süre (saat)", "X,XX"),
    (10, "En zor", "5. çekirdek, pencere dışı TOPLA: taşıma zinciri +1 bit → ortalama süre artışı (ns)", "X,XX"),
]


def tr(x, nd=2):
    return f"{x:,.{nd}f}".replace(",", "X").replace(".", ",").replace("X", ".")


def anahtar_yaz(yol, rows, a, seed, sessiz, olaylar):
    wb = Workbook()
    ws = wb.active
    ws.title = "cevaplar"
    deger = [a["q1"], a["q2"], a["q3"], a["q4"], a["q5"], a["q6"], a["q7"], a["q8"], a["q9"], a["q10"]]
    ints = {1, 3, 4, 5, 6, 8}
    satirlar = []
    for (no, z, soru, bicim), v in zip(SORULAR, deger):
        satirlar.append((no, z, soru, v, tr(v, 0 if no in ints else 2), bicim))
    son = tablo_yaz(ws, ["#", "Zorluk", "Soru", "Cevap", "Cevap (Türkçe biçim)", "Biçim"], satirlar,
                    [5, 12, 78, 16, 20, 15], bicim={4: "#,##0.00"}, hiza={3: "left"},
                    filtre=False, sekme="C00000")
    for i, no in enumerate([s[0] for s in SORULAR], 2):
        ws.cell(i, 4).number_format = "0" if no in ints else "#,##0.00"
        ws.cell(i, 4).font = Font(name=F, size=10, bold=True, color="C00000")
    s_bas, s_bit = sessiz
    ara = [("Seed", seed), ("Satır sayısı", len(rows)),
           ("Termal pencerede işlem", a["q4"]), ("Hatalı kayıt", a["q6"]),
           ("Bozulan bit dağılımı (bit 0 → 7)", " | ".join(str(a["bit_dagilim"].get(i, 0)) for i in range(W))),
           ("Sessiz çekirdek / aralık", f"Çekirdek {SESSIZ}: {zaman(s_bas):%d.%m.%Y %H:%M} – {zaman(s_bit):%d.%m.%Y %H:%M}"),
           ("En uzun 3 bekleme (saat, çekirdek)", " | ".join(f"{g / 3600:.2f} (ç{c})" for g, c in a["gaps"])),
           ("Soru 10: n / R² / ham eğim", f"{a['q10_n']} / {a['q10_r2']:.3f} / {a['q10_ham']:.5f}"),
           ("Motor doğrulaması", "Tüm satırlar ve rastgele testler Python işlemleriyle birebir tutarlı")]
    r = son + 2
    ws.cell(r, 1, "Ara değerler").font = Font(name=F, size=12, bold=True, color="1F3864")
    for k, v in ara:
        r += 1
        ws.merge_cells(start_row=r, start_column=1, end_row=r, end_column=3)
        ws.cell(r, 1, k).font = F_KALIN
        c = ws.cell(r, 4, v)
        c.font, c.alignment = F_NORM, Alignment(horizontal="left")
        ws.merge_cells(start_row=r, start_column=4, end_row=r, end_column=6)

    # formüller
    ws2 = wb.create_sheet("formuller")
    f_satir = [(f"Soru {n}", y, f, nt) for n, y, f, nt in CEVAP_FORMUL]
    f_satir += [(f"Yardımcı {k}", f"{k}2 → {k}5001 ({YARDIMCI_AD[k]})", f.format(r=2),
                 "İlk satır; aşağı doldurulur") for k, f in YARDIMCI.items()]
    son2 = tablo_yaz(ws2, ["Soru / sütun", "Yer", "Formül (İngilizce ad, virgül)", "Not"], f_satir,
                     [14, 34, 110, 70], hiza={2: "left", 3: "left", 4: "left"},
                     duz_metin=(3,), filtre=False, sekme="BF8F00")
    for i in range(2, son2 + 1):
        ws2.cell(i, 3).font = F_MONO
        ws2.cell(i, 4).alignment = Alignment(wrap_text=True, vertical="center")
    ws2.cell(son2 + 2, 1, "Türkçe Excel'de ayraç ';' ve işlev adları Türkçedir (fx penceresinde aratın). "
             "Formüller 5.000 satırlık veriye göre yazıldı; tarih karşılaştırmalarındaki 0,5 sn tolerans bilinçlidir."
             ).font = F_NOT

    # hatalı kayıtlar
    idx = {r["no"]: i for i, r in enumerate(rows, 2)}
    tablo_yaz(wb.create_sheet("hatali_kayitlar"),
              ["İşlem No", "Excel Satırı", "Tarih-Saat", "Çekirdek No", "İşlem Türü", "A (ondalık)",
               "B (ondalık)", "Beklenen (ikilik)", "Günlükteki (ikilik)", "Ters Dönen Bit", "Termal Pencerede"],
              [(r["no"], idx[r["no"]], zaman(r["s"]), r["c"], r["tur"], r["A"], r["B"], ikilik(r["bek"]),
                r["sonuc"], r["poz"], "EVET" if r["pen"] else "HAYIR") for r in a["hatali"]],
              [12, 12, 21, 12, 12, 12, 12, 17, 19, 15, 16], bicim={3: FMT_DT, 8: "@", 9: "@"},
              mono=(8, 9), sekme="C00000")

    # yardımcı sütun değerleri (Python'dan)
    ts_c, onceki, birbir = defaultdict(list), {}, []
    for r in rows:
        onceki_s = onceki.get(r["c"])
        onceki[r["c"]] = r["s"]
        isa = r["A"] - 256 if r["A"] > 127 else r["A"]
        isb = 0 if r["B"] is None else (r["B"] - 256 if r["B"] > 127 else r["B"])
        ts = [r["s"] for r in rows]
        birbir.append([r["no"], isa, isb,
                       1 if r["tur"] == "TOPLA" and not -128 <= isa + isb <= 127 else 0,
                       r["kapi"] * birim(r["s"]) / 100, int(r["pen"]), None, r["bek"],
                       int(int(r["sonuc"], 2) != r["bek"]), r["poz"],
                       None if onceki_s is None else (r["s"] - onceki_s) / 3600])
    ts = [r["s"] for r in rows]
    for i, t in enumerate(ts):
        birbir[i][6] = bisect.bisect_left(ts, t + 3600) - i
    tablo_yaz(wb.create_sheet("yardimci_sutunlar"),
              ["İşlem No"] + [f"{k}: {YARDIMCI_AD[k]}" for k in YARDIMCI], birbir,
              [12] + [17] * 10, bicim={5: "0.00", 11: "0.0000"}, sekme="7F7F7F")
    wb.save(yol)


# ───────────────────────── yazılan dosyadan bağımsız doğrulama ─────────────────────────
def bagimsiz_dogrulama(yol, a):
    try:
        import numpy as np
        import pandas as pd
    except ImportError:
        print("! pandas/numpy yok, bağımsız doğrulama atlandı")
        return
    df = pd.read_excel(yol, sheet_name="islemler", dtype={"A (ikilik)": str, "B (ikilik)": str, "Sonuç (ikilik)": str})
    tar = pd.read_excel(yol, sheet_name="tarife")
    ter = pd.read_excel(yol, sheet_name="termal_olaylar")
    kontrol_et(len(df) == N and list(df.columns) == BASLIKLAR, "okunan tablo yapısı")
    kontrol_et(df["Sonuç (ikilik)"].str.len().eq(8).all(), "ikilik metin uzunluğu")
    df["t"] = df["Tarih-Saat"].dt.round("s")
    sn = (df["t"] - pd.Timestamp(BASLA)).dt.total_seconds().astype("int64").to_numpy()
    kontrol_et(np.all(np.diff(sn) > 0), "zaman sırası / tekillik")
    p = {}
    p["q1"] = int(df["Sonuç (ikilik)"].str.count("1").sum())
    m = pd.merge_asof(df[["t", "Kapı Sayısı"]].sort_values("t"), tar.sort_values("Başlangıç Tarihi"),
                      left_on="t", right_on="Başlangıç Tarihi")
    ekim = m["t"].dt.month == 10
    p["q2"] = round(float((m.loc[ekim, "Kapı Sayısı"] * (m.loc[ekim, "Kapı Başına Enerji (nJ)"] * 100).round()).sum()) / 100, 2)
    is_a = np.where(df["A (ondalık)"] > 127, df["A (ondalık)"] - 256, df["A (ondalık)"])
    is_b = np.where(df["B (ondalık)"].fillna(0) > 127, df["B (ondalık)"].fillna(0) - 256, df["B (ondalık)"].fillna(0))
    top = (df["İşlem Türü"] == "TOPLA").to_numpy()
    p["q3"] = int((top & (((is_a + is_b) < -128) | ((is_a + is_b) > 127))).sum())
    pen = np.zeros(len(df), bool)
    for _, o in ter.iterrows():
        pen |= ((df["Çekirdek No"] == o["Çekirdek No"]) & (df["Tarih-Saat"] >= o["Başlangıç"]) & (df["Tarih-Saat"] <= o["Bitiş"])).to_numpy()
    p["q4"] = int(pen.sum())
    p["q5"] = int((np.searchsorted(sn, sn + 3600, "left") - np.arange(len(sn))).max())
    A_, B_ = df["A (ondalık)"].to_numpy(), df["B (ondalık)"].fillna(0).to_numpy().astype(int)
    A_ = A_.astype(int)
    tur = df["İşlem Türü"].to_numpy()
    bek = np.select([tur == "TOPLA", tur == "ÇIKAR", tur == "AND", tur == "OR", tur == "XOR"],
                    [(A_ + B_) % 256, (A_ - B_) % 256, A_ & B_, A_ | B_, A_ ^ B_], 255 - A_)
    log = df["Sonuç (ikilik)"].map(lambda s: int(s, 2)).to_numpy()
    hat = bek != log
    p["q6"] = int(hat.sum())
    p["q7"] = round(float((hat & pen).sum() / pen.sum() * 100), 2)
    pos = np.log2(bek[hat] ^ log[hat]).round().astype(int)
    p["q8"] = int(np.bincount(pos, minlength=8).argmax())
    gap = df.assign(sn=sn).sort_values(["Çekirdek No", "sn"]).groupby("Çekirdek No")["sn"].diff().max()
    p["q9"] = round(float(gap) / 3600, 2)
    sel = ((df["Çekirdek No"] == 5) & top & ~pen).to_numpy()
    p["q10"] = round(float(np.polyfit(df.loc[sel, "Taşıma Zinciri"], df.loc[sel, "Süre (ns)"], 1)[0]), 2)
    for k, v in p.items():
        kontrol_et(abs(v - a[k]) < 1e-6, f"bağımsız doğrulama uyuşmuyor: {k} → {v} ≠ {a[k]}")
    print("✓ Yazılan xlsx pandas/numpy ile yeniden okunup 10 cevap bağımsız hesaplandı: hepsi tutarlı")


# ───────────────────────────── ana akış ─────────────────────────────
def main():
    cikti = Path(sys.argv[1]) if len(sys.argv) > 1 else Path(__file__).resolve().parent
    cikti.mkdir(parents=True, exist_ok=True)
    for seed in range(SEED0, SEED0 + 500):
        rows, olaylar, sessiz = uret(seed)
        a = cevaplari_hesapla(rows)
        sorun = kontrol(rows, a)
        if not sorun:
            break
        print(f"seed {seed}: {sorun} → yeniden üretiliyor")
    else:
        raise RuntimeError("uygun seed bulunamadı")
    veri, anahtar = cikti / "veri_alu.xlsx", cikti / "cevap_anahtari.xlsx"
    veri_yaz(veri, rows, olaylar)
    anahtar_yaz(anahtar, rows, a, seed, sessiz, olaylar)
    bagimsiz_dogrulama(veri, a)
    print(f"\nseed={seed}  →  {veri.name}, {anahtar.name}  ({cikti})\n")
    for (no, _, soru, _), k in zip(SORULAR, ["q1", "q2", "q3", "q4", "q5", "q6", "q7", "q8", "q9", "q10"]):
        print(f"  S{no:<2} {a[k]:>12}   {soru}")


if __name__ == "__main__":
    main()
