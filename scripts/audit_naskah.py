#!/usr/bin/env python3
"""
audit_naskah.py - Audit read-only naskah ujian (.docx) terhadap standar SD/SMP
SD Swasta Methodist-11 (kertas F4, margin hemat, font Tahoma, kelipatan 0,65 cm,
kop tabel, judul seragam, jarak kolom 1,0 cm, tanpa sisa placeholder).

PASSI: python-docx SAJA (tanpa Word COM, tanpa render) -> cepat & bisa jalan
       di PC mana pun. Opening Word itu berat & bisa nge-hang (LRN-20260918-001).

READ-ONLY: file sumber hanya dibaca. Tidak ada save()/write() ke dokumen.
           (Pola sama rekap_absensi.py LRN-20260928-015 & audit_ujian.py.)

Yang TIDAK bisa dicek tanpa Word/LibreOffice: jumlah halaman sebenarnya,
oposisi baris, dan piksel hasil cetak. Untuk itu pakai fit_1page.py (LRN-20260928-004)
atau audit_scan_pdf.py (LRN-20260929-001).

CLI:
  py audit_naskah.py "naskah.docx" ["naskah2.docx" ...] [--sd|--smp]
  py audit_naskah.py --dir "<folder>" [--recursive] [--skip-backup]
  [--csv Laporan.csv] [--xlsx Laporan.xlsx] [--json] [--strict-placeholder]
  [--font Tahoma] [--min-size 10] [--max-size 14] [--halaman-1]

Exit: 0 bersih, 1 ada temuan, 2 salah argumen.
      (Exit 1 = "berhasil menemukan masalah", bukan tool gagal.)
"""

import sys
import os
import re
import json
import zipfile
import argparse
import collections

try:
    sys.stdout.reconfigure(encoding="utf-8", errors="replace")
    sys.stderr.reconfigure(encoding="utf-8", errors="replace")
except AttributeError:
    pass

from docx import Document
from docx.oxml.ns import qn

EMU_CM = 360000.0
TOL_KERTAS_CM = 0.10
TOL_MARGIN_CM = 0.05
TOL_GRID_CM = 0.012
GRID_CM = 0.65
KOLOM_CM = 1.0

FAIL, WARN, INFO = "FAIL", "WARN", "INFO"

STD = {
    "SD": dict(
        label="SD Swasta Methodist-11",
        kop_re=re.compile(r"SD\s*SWASTA\s*METHODIST", re.I),
        font="Tahoma",
    ),
    "SMP": dict(
        label="SMP Swasta Methodist-11",
        kop_re=re.compile(r"SMP\s*SWASTA\s*METHODIST", re.I),
        font="Tahoma",
    ),
}

# --------------------------------------------------------------- pola
# Selaras dengan ujian_builder.py (NUM_RE/OPT_RE/HDR_BAGIAN_RE) supaya
# deteksi "yang dianggap soal" tidak berbeda antar tool.
SOAL_RE = re.compile(r"^\s*(\d{1,3})[\.\)]\s+(\S.*)$")
OPTI_RE = re.compile(r"^\s*([a-d])[\.\)]\s+(\S.*)$")
BAGIAN_RE = re.compile(r"^\s*[Bb]agian\s+([A-Za-z]|\d|[IVX])[\.\)]?\s+\S")
JUDUL_KATA = re.compile(
    r"^\s*(?:[A-Z][\.\)]\s*(?:C\.?P\.?\s*\d+|soal|pilihan|isian|esai|uraian|essay)"
    r"|(?:isian|esai|uraian|essay|pilihan\s*(?:ber)?ganda|benar\s*[-/]?\s*salah|kunci\s*jawaban)\b"
    r"|(?:C\.?P\.?\s*\d+|PKD)\b)",
    re.I,
)

# Deteksi jenjang untuk mode --auto (folder bisa campuran SD & SMP).
# NAMA kelas muncul dalam DUA urutan di lapangan:
#   digit-depan : "7 SMP", "9 kls", "6 SD"   -> (?:^|\D)([1-9])\s*(kelas|...)
#   kata-depan  : "Kelas 1", "Kls 9"         -> (?:kelas|kls)\s*([1-9])
# HANYA salah satu yang ditambahkan, kasus "Kelas 1" akan TIDAK terdeteksi dan
# file jatuh ke "tidak terdeteksi" padahal isinya jelas-jelas kelas 1 SD.
RE_KELAS_SMP = re.compile(r"(?:^|\D)([789])\s*(?:smp|kls|kelas)\b", re.I)
RE_KELAS_SMP_REV = re.compile(r"\b(?:kelas|kls)\s*([789])\b", re.I)
RE_SMP_KATA = re.compile(r"\bSMP\b", re.I)
# Fallback SD: dipakai HANYA setelah kop SMP / kelas 7-9 gagal, dan kata 'SD'
# sengaja hanya dicocokkan ke NAMA FILE (bukan isi body) supaya kalimat tak
# sengaja (mis. nama jalan, istilah) tidak salah menetapkan jenjang.
RE_KELAS_SD = re.compile(r"(?:^|\D)([1-6])\s*(?:sd|kls|kelas)\b", re.I)
RE_KELAS_SD_REV = re.compile(r"\b(?:kelas|kls)\s*([1-6])\b", re.I)
RE_SD_KATA = re.compile(r"\bSD\b", re.I)

# Placeholder: dipisah dua level keparahan. PISAHKAN sisa template yang pasti
# dari pola yang sah di naskah ujian. "......" adalah places eyed jawaban murid
# di SD -> JANGAN jadi FAIL (jebakan: pola taken dari KSP kebetulan ketemu di
# naskah ujian, bukan sebaliknya).
PLACEHOLDER_FAIL = [
    (re.compile(r"\bUPT\s+SD\b", re.I), "sisa template 'UPT SD'"),
    (re.compile(r"\bSDN\s*/\s*SDS\b", re.I), "sisa template 'SDN / SDS'"),
    (re.compile(r"\(.*uraikan.*\)", re.I), "instruksi kerja '(uraikan ...)'"),
    (re.compile(r"\(.*tambahkan.*\)", re.I), "instruksi kerja '(tambahkan ...)'"),
    (re.compile(r"\(.*contoh.*\)", re.I), "instruksi kerja '(contoh ...)'"),
    (re.compile(r"\[(?:isi|lengkapi|tulis|placeholder|placeholder\s*aja)\]", re.I),
     "kurung siku placeholder [isi/lengkapi/tulis]"),
]
PLACEHOLDER_INFO = [
    (re.compile(r"\.{3,}|\u2026+"), "titik-tik/ellipsis (biasanya kolom jawaban - cek manual)"),
    (re.compile(r"\bLorem\s+ipsum\b", re.I), "teks contoh 'lorem ipsum'"),
    (re.compile(r"\basdf\b", re.I), "teks placeholder 'asdf'"),
]


def cm(v):
    return None if v is None else round(float(v) / EMU_CM, 3)


def twip_cm(v):
    return None if v is None else round(float(v) / 1440.0 * 2.54, 3)


# ------------------------------------------------------- font & size resolution
def _rfonts(rpr):
    if rpr is None:
        return {}
    f = rpr.find(qn("w:rFonts"))
    if f is None:
        return {}
    return {k.split("}")[-1]: v for k, v in f.attrib.items()}


def read_theme_fonts(path):
    """Ambil latin typeface dari word/theme/*.xml (w:asciiTheme -> typeface)."""
    out = {"major": None, "minor": None}
    try:
        with zipfile.ZipFile(path) as z:
            for n in z.namelist():
                if not n.startswith("word/theme/"):
                    continue
                xml = z.read(n).decode("utf-8", "replace")
                for kind in ("major", "minor"):
                    if out[kind] is None:
                        m = re.search(
                            r'<a:%sFont>.*?<a:latin typeface="([^"]*)"' % kind, xml, re.S
                        )
                        if m and m.group(1):
                            out[kind] = m.group(1)
    except Exception:
        pass
    return out


def _resolve_theme_attr(value, theme):
    """asciiTheme/hAnsiTheme: 'minorHAnsi' -> theme minor latin."""
    if not value:
        return None
    v = value.lower()
    if v.startswith("major"):
        return theme.get("major")
    if v.startswith("minor"):
        return theme.get("minor")
    return None


def _style_lookup(doc, key):
    # JEBRAKAN: doc.styles[style_id] emits deprecation warning di python-docx 1.2.
    # Cari by NAMA dulu, baru fallback ke style_id.
    try:
        return doc.styles[key]
    except KeyError:
        pass
    try:
        for s in doc.styles:
            if s.style_id == key:
                return s
    except Exception:
        pass
    return None


def style_chain(doc, style_obj):
    chain, seen = [], set()
    st = style_obj
    while st is not None:
        sid = getattr(st, "style_id", None)
        if sid in seen:
            break
        seen.add(sid)
        chain.append(st)
        st = st.base_style
    return chain


def docdefaults_rpr(doc):
    el = doc.styles.element
    dd = el.find(qn("w:docDefaults"))
    if dd is None:
        return None
    d = dd.find(qn("w:rPrDefault"))
    if d is None:
        return None
    return d.find(qn("w:rPr"))


def _sz_from_rpr(rpr):
    if rpr is None:
        return None
    el = rpr.find(qn("w:sz"))
    if el is None:
        return None
    try:
        return round(int(el.get(qn("w:val"))) / 2.0, 2)
    except (TypeError, ValueError):
        return None


def resolve_font_size(doc, run, para):
    """run rPr -> rStyle -> style chain -> docDefaults. Return pt atau None."""
    cands = []
    rpr = run._r.find(qn("w:rPr"))
    cands.append(rpr)
    if rpr is not None:
        rs = rpr.find(qn("w:rStyle"))
        if rs is not None:
            st = _style_lookup(doc, rs.get(qn("w:val")))
            if st is not None:
                cands.append(st.element.find(qn("w:rPr")))
    try:
        cands.extend(s.element.find(qn("w:rPr")) for s in style_chain(doc, para.style))
    except Exception:
        pass
    cands.append(docdefaults_rpr(doc))
    for c in cands:
        v = _sz_from_rpr(c)
        if v:
            return v
    return None


def resolve_font(doc, run, para, theme):
    """
    Font efektif untuk teks Latin. Urutan: run -> rStyle -> style chain -> docDefaults.
    w:asciiTheme harus DIRESOLVE lewat theme1.xml, kalau tidak font reports "tak ada".
    eastAsia sengaja DIABAIKAN untuk penilaian font (hanya relevan untuk CJK).
    """
    cands = []
    rpr = run._r.find(qn("w:rPr"))
    cands.append(rpr)
    if rpr is not None:
        rs = rpr.find(qn("w:rStyle"))
        if rs is not None:
            st = _style_lookup(doc, rs.get(qn("w:val")))
            if st is not None:
                cands.append(st.element.find(qn("w:rPr")))
    try:
        cands.extend(s.element.find(qn("w:rPr")) for s in style_chain(doc, para.style))
    except Exception:
        pass
    cands.append(docdefaults_rpr(doc))

    east = None
    for c in cands:
        rf = _rfonts(c)
        if not rf:
            continue
        if rf.get("eastAsia"):
            east = east or rf["eastAsia"]
        name = rf.get("ascii") or rf.get("hAnsi")
        if name:
            return name, east
        t = _resolve_theme_attr(rf.get("asciiTheme") or rf.get("hAnsiTheme"), theme)
        if t:
            return t, east
    return None, east


# --------------------------------------------------------------- app.xml pages
def app_pages(path):
    try:
        with zipfile.ZipFile(path) as z:
            for n in z.namelist():
                if n == "docProps/app.xml":
                    m = re.search(rb"<Pages>(\d+)</Pages>", z.read(n))
                    return int(m.group(1)) if m else None
    except Exception:
        return None
    return None


# ------------------------------------------------------------------- audit
def is_kop_table(tbl, kop_re):
    try:
        txt = " ".join(c.text for r in tbl.rows for c in r.cells)
    except Exception:
        return False
    return bool(kop_re.search(re.sub(r"\s+", " ", txt)))


def looks_heading(text):
    """Judul bagian vs opsi a-d.

    JEBRAKAN: pola judul (spesifik) WAJIB dicek lebih dulu daripada pola
    opsi/soal (generik). Kalau dibalik, "A. Pilihan Ganda (100%)" ikut
    cocok OPTI_RE sehingga cek judul-seragam diam-diam tak pernah jalan.
    Opsi juga dibatasi huruf kecil: opsi sekolah 'a. b. c. d.', sedangkan
    judul bagian pakai huruf besar 'A. B. C. D.'.
    """
    if not text or len(text) > 80:
        return False
    if JUDUL_KATA.match(text) or BAGIAN_RE.match(text):
        return True
    if SOAL_RE.match(text) or OPTI_RE.match(text):
        return False
    return False


def _blob_teks(doc, batasi=120000):
    """Gabung paragraf body + seluruh sel tabel (kop ada di tabel)."""
    parts = [p.text for p in doc.paragraphs]
    for t in doc.tables:
        try:
            for row in t.rows:
                for c in row.cells:
                    parts.append(c.text)
        except Exception:
            continue
    return "\n".join(parts)[:batasi]


def deteksi_jenjang(doc, nama_file):
    """Tentukan SD atau SMP per file. Balik (jenjang|None, alasan).

    JEBRAKAN: satu folder soal bisa BERCAMPUR (kelas 1-6 SD dan 7-9 SMP).
    Menjalankan satu preset untuk seluruh folder lalu melaporkan 'kop tidak
    ditemukan' = FALSE POSITIVE pada semua naskah SMP, karena regex preset
    SD memang tidak akan cocok dengan kop 'SMP SWASTA METHODIST-11'.
    Bukti: audit UTS 1 (92 file) melaporkan 35 FAIL kop yang hampir semuanya
    file kelas 7-9. Preset harus ditentukan per FILE, bukan per folder.
    """
    blob = _blob_teks(doc)
    kop_sd = bool(STD["SD"]["kop_re"].search(blob))
    kop_smp = bool(STD["SMP"]["kop_re"].search(blob))
    if kop_smp and not kop_sd:
        return "SMP", "kop 'SMP SWASTA METHODIST-11' terdeteksi"
    if kop_sd and not kop_smp:
        return "SD", "kop 'SD SWASTA METHODIST-11' terdeteksi"
    if kop_smp and kop_sd:
        return None, "kedua kop (SD & SMP) muncul di satu file - ambigu"
    m = None
    for pola in (RE_KELAS_SMP, RE_KELAS_SMP_REV):
        m = pola.search(nama_file) or pola.search(blob)
        if m:
            break
    if m:
        return "SMP", "kelas %s terdeteksi" % m.group(1)
    if RE_SMP_KATA.search(nama_file):
        return "SMP", "keterangan 'SMP' pada nama file"
    m = None
    for pola in (RE_KELAS_SD, RE_KELAS_SD_REV):
        m = pola.search(nama_file) or pola.search(blob)
        if m:
            break
    if m:
        return "SD", "kelas %s terdeteksi" % m.group(1)
    if RE_SD_KATA.search(nama_file):
        return "SD", "keterangan 'SD' pada nama file"
    return None, "tidak terdeteksi"


def audit(path, std="AUTO", font_want=None, min_size=11.0, max_size=14.0,
          strict_placeholder=False):
    res = dict(file=os.path.basename(path), folder=os.path.dirname(path), std=std,
               fail=0, warn=0, info=0, notes=[])
    A = res["notes"].append

    def add(sev, kode, pesan):
        res[sev.lower()] += 1
        res.setdefault("temuan", []).append((sev, kode, pesan))
        A("%-4s %-16s %s" % (sev, kode, pesan))

    try:
        doc = Document(path)
    except Exception as e:
        add(FAIL, "baca", "Gagal dibaca: %s: %s" % (type(e).__name__, e))
        return res

    # Preset per FILE. Default AUTO karena folder soal routinely bercampur
    # kelas 1-6 (SD) dan 7-9 (SMP) dalam satu tree.
    alasan_jenjang = "preset dipaksa operator"
    if std == "AUTO":
        std_efektif, alasan_jenjang = deteksi_jenjang(doc, res["file"])
        if std_efektif is None:
            std_efektif = "SD"
            add(WARN, "jenjang",
                "Jenjang tidak terdeteksi (%s) - dipakai standar SD, "
                "verifikasi manual" % alasan_jenjang)
        else:
            add(INFO, "jenjang", "Standar %s dipakai (%s)" % (std_efektif, alasan_jenjang))
    else:
        std_efektif = std
    res["std"] = std_efektif
    res["std_diminta"] = std
    res["alasan_jenjang"] = alasan_jenjang
    spec = STD[std_efektif]
    font_want = font_want or spec["font"]

    theme = read_theme_fonts(path)

    # --- 1. kertas
    sec = doc.sections[0]
    w, h = cm(sec.page_width), cm(sec.page_height)
    res["kertas"] = "%s x %s cm" % (w, h)
    if abs(w - 21.59) > TOL_KERTAS_CM or abs(h - 33.02) > TOL_KERTAS_CM:
        nama = "A4" if abs(w - 21.0) < 0.3 and abs(h - 29.7) < 0.3 else "?"
        add(FAIL, "kertas",
            "Kertas %s cm, standar F4/Folio 21,59 x 33,02 cm%s"
            % (res["kertas"], " (terbaca A4)" if nama == "A4" else ""))

    # --- 2. margin
    sisi = [("atas", cm(sec.top_margin), 0.5), ("bawah", cm(sec.bottom_margin), 0.6),
            ("kiri", cm(sec.left_margin), 1.0), ("kanan", cm(sec.right_margin), 1.0)]
    res["margin"] = " ".join("%s=%s" % (n, v) for n, v, _ in sisi)
    for n, v, std_cm in sisi:
        if v is None:
            add(WARN, "margin", "Margin %s tidak terbaca" % n)
        elif abs(v - std_cm) > TOL_MARGIN_CM:
            add(FAIL, "margin", "Margin %s %.2f cm, standar %.2f cm" % (n, v, std_cm))

    # --- 3. font efektif + ukuran
    # PENTING: body (doc.paragraphs) DIPISAH dari isi tabel. Kop punya format
    # sendiri (Impact 16pt + Tahoma 7pt) jadi 7pt/16pt di kop itu BENAR, bukan
    # WARN. Kalau dicampur, kop standar malah dilaporkan "ukuran < 10 pt" (false
    # positive). Aturan ketat hanya untuk body = yang diketik guru.
    font_c = collections.Counter()
    size_c = collections.Counter()
    font_t = collections.Counter()
    size_t = collections.Counter()
    east_liar = 0

    def kumpulkan(paras, fc, sc):
        nonlocal east_liar
        for p in paras:
            for r in p.runs:
                if not r.text.strip():
                    continue
                f, east = resolve_font(doc, r, p, theme)
                fc[f or "(tidak terbaca)"] += len(r.text)
                s = resolve_font_size(doc, r, p)
                sc[s if s else "(tidak terbaca)"] += len(r.text)
                if f and east and east != f:
                    east_liar += 1

    kumpulkan(doc.paragraphs, font_c, size_c)
    for tbl in doc.tables:
        for row in tbl.rows:
            for cell in row.cells:
                kumpulkan(cell.paragraphs, font_t, size_t)

    utama = font_c.most_common(1)[0][0] if font_c else "(tidak terbaca)"
    res["font_utama"] = utama
    res["font"] = ", ".join("%s %d%%" % (k, round(100.0 * v / max(1, sum(font_c.values()))))
                            for k, v in font_c.most_common(4))
    if utama == "(tidak terbaca)":
        # Tidak ada satu pun run body dengan font ter-resolve (dokumen kosong, semua
        # isi di textbox/gambar, atau theme tak terbaca). FAIL tanpa bukti = false
        # positive by construction, jadi WARN.
        add(WARN, "font", "Font utama body tidak terbaca - tidak bisa dinilai (periksa manual)")
    elif font_want and font_want.lower() in ("per-soal", "persoal"):
        # Mode per-soal: tiap butir soal wajib seragam secara mandiri (pertanyaan + opsi tidak boleh belang)
        # Boleh beda font antar-soal jika sumber aslinya berbeda.
        belang_soal = []
        cur_blk_fonts = collections.Counter()
        cur_soal_label = "(awal)"
        for p in doc.paragraphs:
            t = p.text.strip()
            if re.search(r'\b(good\s*luck|do\s*your\s*best|have\s*a\s*great|selamat\s*mengerjakan)\b', t, re.I):
                continue
            m = re.match(r'^\s*(\d+|[IVXLCDM]+|[A-Z])[\.\)]\s+', t)
            if m:
                if len(cur_blk_fonts) > 1:
                    belang_soal.append((cur_soal_label, list(cur_blk_fonts.keys())))
                cur_soal_label = m.group(0).strip()
                cur_blk_fonts = collections.Counter()
            for r in p.runs:
                if not r.text.strip() or re.search(r'[\u4e00-\u9fff\u3400-\u4dbf]', r.text):
                    continue
                f, _ = resolve_font(doc, r, p, theme)
                if f:
                    cur_blk_fonts[f] += len(r.text.strip())
        if len(cur_blk_fonts) > 1:
            belang_soal.append((cur_soal_label, list(cur_blk_fonts.keys())))
        if belang_soal:
            rincian = "; ".join("%s: %s" % (lbl, "/".join(fnts)) for lbl, fnts in belang_soal[:3])
            add(WARN, "font", "%d butir soal memiliki font belang/campuran (%s)" % (len(belang_soal), rincian))
        else:
            add(INFO, "font", "Setiap butir soal seragam 100% secara mandiri (mode per-soal)")
    elif font_want and font_want.lower() == "auto":
        # Mode auto: gunakan font dominan dokumen sebagai standar (wajib seragam 100%)
        if len(font_c) > 1:
            campur = ", ".join("%s %d%%" % (k, round(100.0 * v / max(1, sum(font_c.values()))))
                               for k, v in font_c.most_common(4) if k != utama)
            add(WARN, "font", "Font body campuran di luar font utama %s: %s" % (utama, campur))
    elif utama != font_want:
        add(FAIL, "font",
            "Font utama body %s, standar %s (distribusi: %s)" % (utama, font_want, res["font"]))
    elif len(font_c) > 1:
        campur = ", ".join("%s %d%%" % (k, round(100.0 * v / max(1, sum(font_c.values()))))
                           for k, v in font_c.most_common(4) if k != font_want)
        add(WARN, "font", "Font body campuran di luar %s: %s" % (font_want, campur))
    if font_t:
        res["font_tabel"] = ", ".join("%s %d%%" % (k, round(100.0 * v / max(1, sum(font_t.values()))))
                                      for k, v in font_t.most_common(3))
    if east_liar:
        add(INFO, "font", "%d run masih punya w:eastAsia berbeda (tak berpengaruh utk teks Latin)"
            % east_liar)

    kecil = sorted([s for s in size_c if isinstance(s, float) and s < min_size])
    besar = sorted([s for s in size_c if isinstance(s, float) and s > max_size], reverse=True)
    res["ukuran"] = ", ".join("%s pt %d%%" % (k, round(100.0 * v / max(1, sum(size_c.values()))))
                              for k, v in sorted(size_c.items(), key=lambda kv: str(kv[0])))
    for s in kecil:
        add(FAIL, "ukuran-font-floor", "Ukuran body %.2f pt < %.1f pt: melanggar batas bawah font minimal 11.0 pt" % (s, min_size))
    for s in besar:
        add(WARN, "ukuran", "Ukuran body %.2f pt > %.0f pt: risking tidak muat 1 halaman" % (s, max_size))
    if not kecil and not besar and font_c:
        add(INFO, "ukuran", "Ukuran font body %s (lolos batas minimal %.1f pt)" % (res["ukuran"], min_size))
    if size_t:
        res["ukuran_tabel"] = ", ".join(
            "%s pt" % k for k, _ in size_t.most_common(4))
        add(INFO, "ukuran-tabel",
            "Isi tabel tidak dinilai ketat (kop = Impact 16pt + Tahoma 7pt, format sendiri): %s"
            % res["ukuran_tabel"])

    # --- 4. kelipatan indentasi 0,65 cm (body, di luar tabel)
    off, jum = [], 0
    for p in doc.paragraphs:
        pf = p.paragraph_format
        li, fi = cm(pf.left_indent), cm(pf.first_line_indent)
        for v in (li, fi):
            if v is None or abs(v) < 0.01:
                continue
            jum += 1
            k = v / GRID_CM
            if abs(k - round(k)) * GRID_CM > TOL_GRID_CM:
                off.append(v)
    res["indent"] = "%d nilai indent, %d di luar kelipatan" % (jum, len(off))
    if off:
        uniq = sorted(set(off))
        ket = " (kemungkinan default Word 0,75 cm)" if any(
            abs(abs(v) - 0.75) < 0.02 for v in uniq) else ""
        add(FAIL, "indent",
            "%d/%d nilai indent bukan kelipatan 0,65 cm: %s%s"
            % (len(off), jum, ", ".join(str(v) for v in uniq[:6]), ket))

    # --- 5. kop
    kops = [t for t in doc.tables if is_kop_table(t, spec["kop_re"])]
    res["kop"] = len(kops)
    if not kops:
        # Bedakan "tidak bisa dinilai" dari "benar-benar melanggar".
        # Naskah hasil scan/gambar punya 0 tabel TAPI berisi gambar: kop-nya
        # bisa jadi gambar/frame yang tidak terbaca python-docx, jadi SADAR
        # tidak bisa dinilai (WARN). Naskah teks polos (0 tabel DAN 0 gambar)
        # memang benar-benar tidak punya kop -> FAIL.
        n_draw = doc.element.xml.count("<w:drawing>")
        if not doc.tables and n_draw:
            add(WARN, "kop",
                "Dokumen 0 tabel tapi %d gambar - kop mungkin gambar/frame, "
                "tidak bisa dinilai otomatis; periksa manual" % n_draw)
        elif not doc.tables:
            add(FAIL, "kop",
                "Dokumen tanpa tabel sama (0 tabel, 0 gambar) - tidak ada kop")
        else:
            add(FAIL, "kop",
                "%d tabel tapi tak ada yang cocok '%s' (kemungkinan naskah "
                "berjenjang lain atau kop hilang)" % (len(doc.tables), spec["label"]))
    elif len(kops) > 1:
        add(INFO, "kop", "%d kop ditemukan (boleh utk format 2-in-1 v1.6)" % len(kops))
    else:
        add(INFO, "kop", "1 kop ditemukan (wajar)")

    # --- 6. judul seragam & tanpa auto-numbering
    judul, anum = [], 0
    for p in doc.paragraphs:
        t = p.text.strip()
        if not looks_heading(t):
            continue
        bold = any(r.bold for r in p.runs if r.text.strip())
        und = any(r.underline for r in p.runs if r.text.strip())
        judul.append((t[:48], bold, und))
        pPr = p._p.find(qn("w:pPr"))
        if pPr is not None and pPr.find(qn("w:numPr")) is not None:
            anum += 1
    res["judul"] = len(judul)
    if judul:
        gaya = collections.Counter((b, u) for _, b, u in judul)
        if len(gaya) > 1:
            maior = gaya.most_common(1)[0][0]
            beda = [t for t, b, u in judul if (b, u) != maior]
            add(FAIL, "judul-seragam",
                "Format judul tidak seragam. Mayoritas bold=%s underline=%s, tetapi: %s"
                % (maior[0], maior[1], "; ".join(beda[:4])))
        else:
            add(INFO, "judul-seragam", "%d judul, format seragam (bold=%s underline=%s)"
                % (len(judul), gaya.most_common(1)[0][0][0], gaya.most_common(1)[0][0][1]))
    if anum:
        add(FAIL, "judul-numpr",
            "%d judul pakai auto-numbering (w:numPr) - standar terbaru: judul harus teks murni"
            % anum)

    # --- 6b. Jeda vertikal judul (Zero-gap Heading 1 & Jeda Seksi Heading 2+)
    headings_found = []
    prev_p = None
    for p in doc.paragraphs:
        t = p.text.strip()
        if looks_heading(t):
            pf = p.paragraph_format
            sb = pf.space_before.pt if pf.space_before is not None else 0.0
            prev_blank = (prev_p is not None and not prev_p.text.strip())
            headings_found.append((t, sb, prev_blank))
        prev_p = p

    for h_idx, (ht, h_sb, h_blank) in enumerate(headings_found):
        # Heading 1 (awal dokumen atau awal paruh kedua 2-in-1): wajib zero-gap rapat di bawah kop (sb <= 1.5 pt)
        if h_idx == 0 or (len(kops) > 1 and h_idx == len(headings_found) // 2):
            if h_sb > 1.5:
                add(FAIL, "judul-awal-jarak",
                    "Judul pertama '%s' memiliki space_before %.1f pt (wajib rapat 0 pt di bawah kop)"
                    % (ht[:30], h_sb))
            else:
                add(INFO, "judul-awal-jarak", "Judul pertama '%s' rapat 0 pt di bawah kop" % ht[:30])
        else:
            # Heading 2+ (CP 2, Essay, dsb): wajib ada jeda seksi tegas (sb >= 2.0 pt atau ada enter kosong di atasnya)
            if h_sb < 2.0 and not h_blank:
                add(FAIL, "judul-jeda-seksi",
                    "Judul seksi '%s' menempel ke soal di atasnya tanpa jeda seksi (space_before wajib >= 2 pt)"
                    % ht[:30])
            else:
                effective_gap = h_sb if h_sb >= 2.0 else (12.0 if h_blank else 0)
                add(INFO, "judul-jeda-seksi", "Judul seksi '%s' memiliki jeda seksi tegas (gap=%.1f pt)" % (ht[:30], effective_gap))

    # --- 6c. Pemisah Simetris Garis Potong 16.50 cm pada Format 2-in-1
    if len(kops) > 1:
        tbl2_elem = kops[1]._tbl
        prev_elem = tbl2_elem.getprevious()
        has_cut_gap = False
        cut_desc = "tidak ada"
        if prev_elem is not None:
            if prev_elem.tag.endswith('p'):
                p_text = prev_elem.text or ""
                pPr = prev_elem.find(qn('w:pPr'))
                sb_val = 0
                if pPr is not None:
                    sp = pPr.find(qn('w:spacing'))
                    if sp is not None and sp.get(qn('w:before')):
                        sb_val = int(sp.get(qn('w:before'))) / 20.0
                if not p_text.strip() or sb_val >= 6.0:
                    has_cut_gap = True
                    cut_desc = "enter pemisah / gap %.1f pt" % (sb_val if sb_val > 0 else 12.0)
            elif prev_elem.tag.endswith('tbl'):
                has_cut_gap = False
        
        if not has_cut_gap:
            add(FAIL, "2in1-pemisah-tengah",
                "Kop kedua (2-in-1) menempel langsung ke soal di atasnya tanpa jeda garis potong simetris 16.50 cm (wajib enter pemisah)")
        else:
            add(INFO, "2in1-pemisah-tengah",
                "Kop kedua (2-in-1) memiliki jeda garis potong simetris 16.50 cm (%s)" % cut_desc)

    # --- 7. jarak kolom
    cols = sec._sectPr.find(qn("w:cols"))
    if cols is not None:
        n = int(cols.get(qn("w:num")) or 1)
        sp = cols.get(qn("w:space"))
        spc = twip_cm(int(sp)) if sp and sp.isdigit() else None
        res["kolom"] = "%d kolom spasi=%s cm" % (n, spc)
        if n >= 2 and (spc is None or abs(spc - KOLOM_CM) > 0.06):
            add(FAIL, "kolom",
                "%d kolom dengan spasi %s cm, standar 1,0 cm (0,65 cm memotong nomor 2 digit)"
                % (n, spc))

    # --- 8. nomor soal (literal vs auto) - info pipeline
    lit = sum(1 for p in doc.paragraphs if SOAL_RE.match(p.text.strip()))
    auto = 0
    for p in doc.paragraphs:
        pPr = p._p.find(qn("w:pPr"))
        if pPr is not None and pPr.find(qn("w:numPr")) is not None and not looks_heading(p.text.strip()):
            auto += 1
    res["nomor"] = "literal=%d auto=%d" % (lit, auto)

    # --- 9. placeholder
    tails = []
    for p in doc.paragraphs:
        tails.append(p.text.strip())
    for tbl in doc.tables:
        try:
            for row in tbl.rows:
                for c in row.cells:
                    tails.append(c.text.strip())
        except Exception:
            pass
    lvl = PLACEHOLDER_FAIL if strict_placeholder else PLACEHOLDER_FAIL
    for rx, ket in lvl:
        for t in tails:
            if t and rx.search(t):
                add(FAIL, "placeholder", "%s: \"%s\"" % (ket, t[:70]))
                break
    for rx, ket in PLACEHOLDER_INFO:
        for t in tails:
            if t and rx.search(t):
                add(INFO, "placeholder", "%s" % ket)
                break

    # --- 10. jumlah halaman (CACHE!)
    pg = app_pages(path)
    res["halaman_cache"] = pg
    if pg is None:
        add(INFO, "halaman", "docProps/app.xml tidak punya <Pages> - jumlah halaman tak bisa "
                             "dibaca tanpa Word/LibreOffice")
    else:
        add(INFO, "halaman", "app.xml Pages=%d (CACHE hasil save terakhir, bisa basi) "
                             "- verifikasi dgn fit_1page.py / LibreOffice" % pg)

    res["ringkas"] = "%s | kertas %s | font %s | FAIL %d WARN %d" % (
        res["file"], res["kertas"], utama, res["fail"], res["warn"])
    return res


# ------------------------------------------------------------------- IO
def collect(paths, folder, recursive, skip_backup):
    files = []
    for p in paths:
        if os.path.isdir(p):
            files.extend(scan_dir(p, recursive, skip_backup))
        elif os.path.isfile(p):
            files.append(p)
        else:
            sys.stderr.write("[!] Path tidak ditemukan: %s\n" % p)
    return files


def scan_dir(folder, recursive, skip_backup):
    out = []
    if recursive:
        for root, dirs, names in os.walk(folder):
            dirs[:] = [d for d in dirs if not (skip_backup and d.lower() in ("@backup", "backup"))]
            for n in names:
                if n.lower().endswith(".docx") and not n.startswith("~"):
                    out.append(os.path.join(root, n))
    else:
        try:
            for n in os.listdir(folder):
                p = os.path.join(folder, n)
                if os.path.isfile(p) and n.lower().endswith(".docx") and not n.startswith("~"):
                    out.append(p)
        except OSError as e:
            sys.stderr.write("[!] %s\n" % e)
    return sorted(out)


def export_csv(rows, out):
    import csv
    cols = ["file", "folder", "std", "kertas", "margin", "font", "ukuran", "indent",
            "kop", "judul", "kolom", "nomor", "halaman_cache", "fail", "warn", "info",
            "catatan"]
    with open(out, "w", newline="", encoding="utf-8-sig") as f:
        w = csv.DictWriter(f, fieldnames=cols)
        w.writeheader()
        for r in rows:
            row = {c: r.get(c, "") for c in cols}
            row["catatan"] = " | ".join(r["notes"])
            w.writerow(row)


def export_xlsx(rows, out):
    from openpyxl import Workbook
    from openpyxl.styles import Font, PatternFill, Alignment
    from openpyxl.utils import get_column_letter
    from openpyxl.worksheet.properties import PageSetupProperties

    cols = ["file", "std", "kertas", "margin", "font", "ukuran", "indent", "kop",
            "judul", "kolom", "nomor", "halaman_cache", "fail", "warn", "info", "catatan"]
    wb = Workbook()
    ws = wb.active
    ws.title = "Audit Naskah"
    ws.append(cols)
    for c in ws[1]:
        c.font = Font(bold=True, color="FFFFFF")
        c.fill = PatternFill("solid", fgColor="1F2A57")
        c.alignment = Alignment(horizontal="center", vertical="center")
    for r in rows:
        ws.append([r.get(c, "") for c in cols])
    hf = PatternFill("solid", fgColor="FFC7CE")
    for i, r in enumerate(rows, start=2):
        if r["fail"]:
            ws.cell(i, cols.index("fail") + 1).fill = hf
    ws.freeze_panes = "A2"
    ws.auto_filter.ref = ws.dimensions
    for i, c in enumerate(cols, start=1):
        w = 46 if c in ("catatan", "font", "ukuran", "margin") else max(12, min(28, len(c) + 4))
        ws.column_dimensions[get_column_letter(i)].width = w
    ws.sheet_properties.pageSetUpPr = PageSetupProperties(fitToPage=True)
    ws.page_setup.orientation = "landscape"
    ws.page_setup.fitToWidth = 1
    ws.page_setup.fitToHeight = 0
    ws.print_title_rows = "1:1"
    last = get_column_letter(len(cols))
    ws.print_area = "A1:%s%d" % (last, ws.max_row)
    wb.save(out)


def main(argv=None):
    ap = argparse.ArgumentParser(description="Audit read-only naskah ujian .docx (standar Methodist-11)")
    ap.add_argument("paths", nargs="*", help="file .docx atau folder")
    ap.add_argument("--dir", help="folder yang mau diaudit")
    ap.add_argument("--recursive", action="store_true", help="ikut subfolder")
    ap.add_argument("--skip-backup", action="store_true", help="lewati folder @backup/backup")
    ap.add_argument("--sd", action="store_true", help="paksa standar SD")
    ap.add_argument("--smp", action="store_true", help="paksa standar SMP")
    ap.add_argument("--auto", action="store_true",
                    help="deteksi SD/SMP per file (default)")
    ap.add_argument("--font", help="override nama font standar")
    ap.add_argument("--min-size", type=float, default=11.0)
    ap.add_argument("--max-size", type=float, default=14.0)
    ap.add_argument("--strict-placeholder", action="store_true")
    ap.add_argument("--csv")
    ap.add_argument("--xlsx")
    ap.add_argument("--json", action="store_true")
    ap.add_argument("--halaman-1", action="store_true", help="peringatan jika app.xml Pages > 1")
    args = ap.parse_args(argv)

    paths = list(args.paths)
    if args.dir:
        paths.append(args.dir)
    if not paths:
        ap.error("beri minimal satu file .docx atau --dir <folder>")

    if args.sd and args.smp:
        ap.error("--sd dan --smp tidak bisa dipakai bersamaan")
    std = "AUTO"
    if args.sd:
        std = "SD"
    elif args.smp:
        std = "SMP"
    files = collect(paths, args.dir, args.recursive, args.skip_backup)
    # jebakan: lock file Word '~$' bukan naskah
    files = [f for f in files if not os.path.basename(f).startswith("~")]
    files = [f for f in files if f.lower().endswith((".docx", ".doc"))]
    if not files:
        sys.stderr.write("[!] Tidak ada file .docx ditemukan\n")
        return 2

    info = sys.stderr if args.json else sys.stdout
    label = "AUTO (SD/SMP per file)" if std == "AUTO" else STD[std]["label"]
    print("[*] Standar: %s | %d file (READ-ONLY, tanpa Word COM)" % (label, len(files)),
          file=info)

    rows = []
    for f in files:
        if f.lower().endswith(".doc"):
            print("[!] Lewati %s (format .doc lama; konversi ke .docx dulu)" % os.path.basename(f),
                  file=info)
            continue
        r = audit(f, std, args.font, args.min_size, args.max_size, args.strict_placeholder)
        if args.halaman_1 and isinstance(r.get("halaman_cache"), int) and r["halaman_cache"] > 1:
            r["notes"].append("INFO   halaman          app.xml Pages=%d > 1 (target: pas 1 halaman)"
                              % r["halaman_cache"])
            r["info"] += 1
        rows.append(r)
        print("\n--- %s" % r["file"], file=info)
        for n in r["notes"]:
            print("  %s" % n, file=info)

    tf = sum(r["fail"] for r in rows)
    tw = sum(r["warn"] for r in rows)
    ti = sum(r["info"] for r in rows)
    print("\n%s" % ("=" * 78), file=info)
    print("RINGKASAN: %d file | FAIL %d | WARN %d | INFO %d" % (len(rows), tf, tw, ti), file=info)
    if tf == 0:
        print("[+] Tidak ada pelanggaran standar.", file=info)
    else:
        print("[!] %d pelanggaran FAIL - perbaiki sebelum masuk tahap OK/P." % tf, file=info)

    # Distribusi WAJIB ditampilkan: jumlah FAIL bukan jumlah file rusak. Satu file
    # bisa punya 4 margin sekaligus, dan 1 kode bisa mendominasi seluruh korpus
    # karena Mixing-era/preset, bukan karena 80 file benar-benar rusak (lihat
    # LRN-20260929-004: anomali berfrekuensi tinggi = konvensi, bukan cacat).
    kode = collections.Counter()
    for r in rows:
        # row in-memory punya "notes" (list str); "catatan" hanya ada di row ekspor.
        for p in [x.strip() for x in (r.get("notes") or []) if x and x.strip()]:
            bagian = p.split(None, 2)
            if len(bagian) >= 2 and bagian[0] in (FAIL, WARN):
                kode["%s %s" % (bagian[0], bagian[1])] += 1
    if kode:
        print("\nDistribusi temuan per kode (1 file bisa dihitung >1 kali):", file=info)
        for k, v in kode.most_common():
            print("   %4d  %s" % (v, k), file=info)
    bersih = sum(1 for r in rows if r["fail"] == 0)
    if bersih:
        print("\n%d dari %d file tanpa FAIL sama sekali." % (bersih, len(rows)), file=info)

    # Rincian per jenjang: folder soal biasanya campuran kelas 1-6 & 7-9.
    # Tanpa ini, FAIL marginalized bisa disalaharti sebagai satu masalah tunggal.
    per_jenjang = collections.defaultdict(lambda: dict(file=0, fail=0, bersih=0))
    for r in rows:
        j = per_jenjang[r.get("std", "?")]
        j["file"] += 1
        j["fail"] += r["fail"]
        if r["fail"] == 0:
            j["bersih"] += 1
    if len(per_jenjang) > 1:
        print("\nRincian per jenjang (folder campuran):", file=info)
        for k in sorted(per_jenjang):
            j = per_jenjang[k]
            print("   %-4s %3d file | FAIL %3d | tanpa FAIL %3d"
                  % (k, j["file"], j["fail"], j["bersih"]), file=info)
    ambigu = [r["file"] for r in rows
              if r.get("std_diminta") == "AUTO" and r.get("alasan_jenjang") == "tidak terdeteksi"]
    if ambigu:
        print("\n[WARN] %d file jenjangnya tidak terdeteksi (dianggap SD): %s"
              % (len(ambigu), ", ".join(ambigu[:6])), file=info)

    if args.json:
        print(json.dumps(rows, ensure_ascii=False, indent=2, default=str))
    if args.csv:
        export_csv(rows, args.csv)
        print("[OK] CSV -> %s" % args.csv, file=info)
    if args.xlsx:
        export_xlsx(rows, args.xlsx)
        print("[OK] XLSX -> %s" % args.xlsx, file=info)
    return 1 if tf else 0


if __name__ == "__main__":
    sys.exit(main())
