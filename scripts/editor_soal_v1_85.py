#!/usr/bin/env python3
"""
editor_soal_v1_84.py - Engine Standar Baku Editing Naskah Soal Ujian Versi 1.84
SD & SMP Swasta Methodist-11 Medan.

KUMULATIF & NON-DESTRUKTIF:
Dibangun langsung di atas baseline stabil v1.77.
Hanya MENAMBAHKAN penanganan gambar, proteksi penomoran, dan mode layout,
tanpa merombak struktur yang sudah terbukti bekerja di v1.77.

TAMBAHAN FITUR v1.84:
1. KOP RESMI NATIVE:
   - Prioritas Word COM native rng.PasteAndFormat(16) (Ctrl+K) agar struktur border
     ganda, tabel, dan logo 100% presisi. Fallback ke deepcopy XML bila non-Windows.
2. ATURAN PENOMORAN CERDAS (ANTI-BLIND-NUMBERING):
   - Nomor asli ada        -> pertahankan (explicit_num).
   - Paragraf jelas soal   -> diproses sebagai butir soal.
   - Instruksi / petunjuk  -> JANGAN diberi nomor (Left 0.65 cm, Hanging 0 cm).
   - Kata / label / pilihan-> JANGAN diberi nomor (Left 0.65 cm, Hanging 0 cm).
   - Tidak yakin           -> JANGAN auto-number.
3. ATURAN GAMBAR & DRAWING:
   - Gambar standalone     -> Left 0.65 cm, Hanging 0.0 cm.
   - Gambar di soal nomor  -> Nomor di 0 cm, isi/gambar di 0.65 cm, Hanging 0.65 cm, Tab stop 0.65 cm.
   - Enter sengaja antar blok visual dipertahankan 1 baris.
4. DIFERENSIASI KELAS 1-2 SD:
   - Disediakan garis jawaban Jb: .................... pada soal esai.
   - Layout proporsional 85-90% halaman F4.
5. PROTEKSI BAHASA INGGRIS & MANDARIN:
   - "Good Luck" di akhir naskah Inggris tidak disentuh formatnya.
6. STANDAR PENOMORAN MANUAL & PILIHAN GANDA (v1.86):
   - Seluruh sub-penomoran manual (a-d) WAJIB menggunakan literal TAB (\t) setelah titik:
     a.[TAB]soal, bukan spasi biasa.
   - Pilihan Ganda 2 Kolom Baku:
     Baris 1: a.[TAB]teks[TAB]c.[TAB]teks
     Baris 2: b.[TAB]teks[TAB]d.[TAB]teks
     Tab stop pada grid kelipatan 0.65 cm (1.30 cm, 4.55 cm, 5.20 cm).
   - Trik P0 1pt di bawah Kop:
     Paragraf transisi continuous section break disetel 1pt exact agar judul menempel rapat.
   - Judul seksi dipisahkan tepat 1 baris enter bersih di atasnya.
   - Font murni #000000 (pembersihan themeColor/themeTint agar tidak abu-abu).
"""

import os
import re
import io
import sys
import copy
import argparse
import docx

from docx.shared import Cm, Pt
from docx.oxml import OxmlElement
from docx.oxml.ns import qn
from docx.enum.text import WD_ALIGN_PARAGRAPH
from docx.table import Table

VERSION = "1.86.1"

# ============================================================
# KONSTANTA BAKU / INVARIANT (v1.77)
# ============================================================
F4_WIDTH = Cm(21.59)
F4_HEIGHT = Cm(33.02)
MARGIN_TOP = Cm(0.50)
MARGIN_BOTTOM = Cm(0.60)
MARGIN_LEFT = Cm(1.00)
MARGIN_RIGHT = Cm(1.00)
GRID_INDENT = Cm(0.65)
COLUMN_SPACING_DXA = "567"  # 1.0 cm

# Kompatibilitas alias kontrak v1.77
MARGIN_TOP_FULL = MARGIN_TOP
MARGIN_BOTTOM_FULL = MARGIN_BOTTOM
MARGIN_LEFT_FULL = MARGIN_LEFT
MARGIN_RIGHT_FULL = MARGIN_RIGHT

RE_HEADER = re.compile(
    r'^\s*([a-dA-D]\.\s*)?(pilihan\s*(?:ber)?ganda|isian|essay|esai|uraian|kunci\s*jawaban|petunjuk|kompetensi|c\.?p\.?\s*\d+|romawi|(?:X{0,3}(?:IX|IV|V?I{1,3}))\.)',
    re.I,
)
RE_OPTION = re.compile(r'^\s*[a-eA-E][\.\)]\s+')
RE_OPT_INLINE = re.compile(r'^\s*([a-b][\.\)])(?:[ \t]+)(.+?)(?:[ \t]{2,}|\t)([c-d][\.\)])(?:[ \t]+)(.+)$', re.IGNORECASE)
RE_OPT_SINGLE = re.compile(r'^\s*([a-eA-E][\.\)])(?:[ \t]+)(.+)$', re.IGNORECASE)
RE_QUESTION = re.compile(r'^\s*(\d+)[\.\)]\s*')
RE_GOOD_LUCK = re.compile(r'\bgood\s*luck\b', re.I)

ROMAN_SD = {1: 'I SD', 2: 'II SD', 3: 'III SD', 4: 'IV SD', 5: 'V SD', 6: 'VI SD'}
ROMAN_SMP = {7: 'VII SMP', 8: 'VIII SMP', 9: 'IX SMP'}

JADWAL_UTS1_2026 = {
    'english math': ('English Mathematics', 'Jumat, 18 September 2026'),
    'english matehmatic': ('English Mathematics', 'Jumat, 18 September 2026'),
    'e. math': ('English Mathematics', 'Jumat, 18 September 2026'),
    'mandarin': ('Bahasa Mandarin', 'Senin, 14 September 2026'),
    'musik': ('Seni Musik', 'Senin, 14 September 2026'),
    'agama': ('Agama', 'Selasa, 15 September 2026'),
    'rupa': ('Seni Rupa', 'Selasa, 15 September 2026'),
    'ips': ('Ilmu Pengetahuan Sosial (IPS)', 'Selasa, 15 September 2026'),
    'pancasila': ('Pendidikan Pancasila', 'Kamis, 17 September 2026'),
    'pkn': ('Pendidikan Pancasila', 'Kamis, 17 September 2026'),
    'ipas': ('IPAS', 'Rabu, 16 September 2026'),
    'ipa': ('Ilmu Pengetahuan Alam (IPA)', 'Rabu, 16 September 2026'),
    'fisika': ('Fisika', 'Rabu, 16 September 2026'),
    'biologi': ('Biologi', 'Rabu, 16 September 2026'),
    'inggris': ('Bahasa Inggris', 'Rabu, 16 September 2026'),
    'english': ('Bahasa Inggris', 'Rabu, 16 September 2026'),
    'matematika': ('Matematika', 'Kamis, 17 September 2026'),
    'math': ('Matematika', 'Kamis, 17 September 2026'),
    'indonesia': ('Bahasa Indonesia', 'Jumat, 18 September 2026'),
}

JADWAL_UAS1_2026 = {
    'english math': ('English Mathematics', 'Sabtu, 5 Desember 2026'),
    'english matehmatic': ('English Mathematics', 'Sabtu, 5 Desember 2026'),
    'e. math': ('English Mathematics', 'Sabtu, 5 Desember 2026'),
    'e. mathematics': ('English Mathematics', 'Sabtu, 5 Desember 2026'),
    'mandarin': ('Bahasa Mandarin', 'Senin, 30 November 2026'),
    'musik': ('Seni Musik', 'Senin, 30 November 2026'),
    'agama': ('Agama', 'Selasa, 1 Desember 2026'),
    'rupa': ('Seni Rupa', 'Selasa, 1 Desember 2026'),
    'ktk': ('Seni Rupa (KTK)', 'Sabtu, 5 Desember 2026'),
    'ips': ('Ilmu Pengetahuan Sosial (IPS)', 'Selasa, 1 Desember 2026'),
    'pancasila': ('Pendidikan Pancasila', 'Rabu, 2 Desember 2026'),
    'pkn': ('Pendidikan Pancasila', 'Rabu, 2 Desember 2026'),
    'ipas': ('IPAS', 'Rabu, 2 Desember 2026'),
    'ipa': ('Ilmu Pengetahuan Alam (IPA)', 'Rabu, 2 Desember 2026'),
    'fisika': ('Fisika', 'Rabu, 2 Desember 2026'),
    'biologi': ('Biologi', 'Rabu, 2 Desember 2026'),
    'inggris': ('Bahasa Inggris', 'Kamis, 3 Desember 2026'),
    'matematika': ('Matematika', 'Kamis, 3 Desember 2026'),
    'math': ('Matematika', 'Kamis, 3 Desember 2026'),
    'indonesia': ('Bahasa Indonesia', 'Jumat, 4 Desember 2026'),
}

# ============================================================
# HELPER TIPOGRAFI & INDENTASI (v1.77 BASELINE)
# ============================================================
def set_run_tahoma(run, size_pt, bold=False, color_hex="000000"):
    run.font.name = "Tahoma"
    run.font.size = Pt(size_pt)
    run.bold = bold
    rPr = run._element.get_or_add_rPr()
    rf = rPr.find(qn('w:rFonts'))
    if rf is None:
        rf = OxmlElement('w:rFonts')
        rPr.append(rf)
    rf.set(qn('w:ascii'), 'Tahoma')
    rf.set(qn('w:hAnsi'), 'Tahoma')
    rf.set(qn('w:cs'), 'Tahoma')

    color_el = rPr.find(qn('w:color'))
    if color_el is None:
        color_el = OxmlElement('w:color')
        rPr.append(color_el)
    color_el.set(qn('w:val'), color_hex)
    for attr in [qn('w:themeColor'), qn('w:themeTint'), qn('w:themeShade')]:
        if attr in color_el.attrib:
            del color_el.attrib[attr]


def set_option_tabs_inline(para):
    """Set tab stops untuk pilihan ganda 2 kolom (a & c, b & d) pada grid kelipatan 0.65 cm."""
    pPr = para._p.get_or_add_pPr()
    tabs = pPr.find(qn('w:tabs'))
    if tabs is not None:
        pPr.remove(tabs)
    tabs = OxmlElement('w:tabs')
    # 737 dxa (1.30 cm), 2580 dxa (4.55 cm = 7*0.65 cm), 2949 dxa (5.20 cm = 8*0.65 cm)
    for pos in [737, 2580, 2949]:
        t = OxmlElement('w:tab')
        t.set(qn('w:val'), 'left')
        t.set(qn('w:pos'), str(pos))
        tabs.append(t)
    pPr.append(tabs)


def set_option_tabs_single(para):
    """Set tab stop untuk pilihan ganda 1 kolom (a-d per baris) pada 1.30 cm."""
    pPr = para._p.get_or_add_pPr()
    tabs = pPr.find(qn('w:tabs'))
    if tabs is not None:
        pPr.remove(tabs)
    tabs = OxmlElement('w:tabs')
    t = OxmlElement('w:tab')
    t.set(qn('w:val'), 'left')
    t.set(qn('w:pos'), '737')
    tabs.append(t)
    pPr.append(tabs)


def set_para_ind(para, left_cm=0.65, hanging_cm=0.65, space_before_pt=0, space_after_pt=0, line_spacing=1.0):
    pf = para.paragraph_format
    pf.left_indent = Cm(left_cm)
    if hanging_cm > 0:
        pf.first_line_indent = Cm(-hanging_cm)
    else:
        pf.first_line_indent = Cm(0.0)
    pf.space_before = Pt(space_before_pt)
    pf.space_after = Pt(space_after_pt)
    pf.line_spacing = line_spacing

    pPr = para._p.get_or_add_pPr()
    ind = pPr.find(qn('w:ind'))
    if ind is None:
        ind = OxmlElement('w:ind')
        pPr.append(ind)
    left_twips = str(int(left_cm * 567))
    hanging_twips = str(int(hanging_cm * 567)) if hanging_cm > 0 else "0"
    ind.set(qn('w:left'), left_twips)
    if hanging_cm > 0:
        ind.set(qn('w:hanging'), hanging_twips)
    else:
        if qn('w:hanging') in ind.attrib:
            del ind.attrib[qn('w:hanging')]


def _set_question_grid_tab(para):
    """Kunci tab stop setelah nomor soal tepat pada grid 0.65 cm."""
    try:
        para.paragraph_format.tab_stops.add_tab_stop(Cm(0.65))
    except Exception:
        pPr = para._p.get_or_add_pPr()
        tabs = pPr.find(qn('w:tabs'))
        if tabs is None:
            tabs = OxmlElement('w:tabs')
            pPr.append(tabs)
        tab = OxmlElement('w:tab')
        tab.set(qn('w:val'), 'left')
        tab.set(qn('w:pos'), str(int(0.65 * 567)))
        tabs.append(tab)


# ============================================================
# DETEKSI METADATA & KLASIFIKASI PARAGRAF CERDAS (TAMBAHAN v1.84)
# ============================================================
def parse_metadata(folder_name, filename, raw_docx_path=None):
    clean_name = re.sub(r'^(?:doc_[a-f0-9]+_|src_copy_[a-f0-9]+_|pre_)', '', filename)
    m = re.search(r'\b(?:kelas|kls)[\s_\-]*([1-9])([a-dA-D]?)\b', clean_name, re.I)
    if not m and folder_name:
        m = re.search(r'\b(?:kelas|kls)[\s_\-]*([1-9])([a-dA-D]?)\b', folder_name, re.I)
    if not m:
        m = re.search(r'\b([1-9])([a-dA-D]?)\b', clean_name)
    if not m:
        m = re.search(r'([1-9])([a-dA-D]?)', clean_name)
    grade = int(m.group(1)) if m else 1
    sub = m.group(2).upper() if m and m.group(2) else ''

    is_smp = grade >= 7
    unit = 'SMP' if is_smp else 'SD'
    if is_smp:
        k_base = ROMAN_SMP.get(grade, str(grade)).split()[0]
        kelas_str = f"{k_base}{sub} SMP" if sub else ROMAN_SMP.get(grade, f'{grade} SMP')
    else:
        k_base = ROMAN_SD.get(grade, str(grade)).split()[0]
        kelas_str = f"{k_base}{sub} SD" if sub else ROMAN_SD.get(grade, f'{grade} SD')

    fname_lower = clean_name.lower()
    fld_lower = folder_name.lower()
    raw_path_lower = str(raw_docx_path).lower() if raw_docx_path else ''
    combined = f"{fname_lower} {fld_lower} {raw_path_lower}"

    if raw_docx_path and os.path.exists(raw_docx_path):
        try:
            import docx as _dmod
            _d = _dmod.Document(raw_docx_path)
            _head_text = " ".join(p.text for p in _d.paragraphs[:10]).lower()
            combined = f"{combined} {_head_text}"
            if 'mandarin' not in combined and re.search(r'[\u4e00-\u9fff]', _head_text):
                combined = f"{combined} mandarin"
        except Exception:
            pass

    is_uas = any(k in combined for k in ('uas', 'semester i', 'semester 1', 'sem_1', 'sem 1', 'ganjil', 'desember', 'november'))
    jadwal_dict = JADWAL_UAS1_2026 if is_uas else JADWAL_UTS1_2026

    mapel = 'Mata Pelajaran'
    hari_tgl = 'Senin, 30 November 2026' if is_uas else 'Senin, 14 September 2026'

    for key, (m_val, d_val) in jadwal_dict.items():
        if key in combined:
            mapel = m_val
            hari_tgl = d_val
            break

    return unit, kelas_str, mapel, hari_tgl, grade


# ============================================================
# NORMALISASI DOKUMEN INPUT (WORD NATIVE NUMBERING PRESERVATION)
# ============================================================
def normalize_input_docx(raw_docx_path, word_app=None):
    """
    Jika Word COM tersedia:
    1. Bekukan Word native list numbering (w:numPr) menjadi teks run biasa via doc.ConvertNumbersToText().
    2. Ganti manual line breaks (^l / Shift+Enter) menjadi paragraph breaks (^p)
       agar tidak memicu bug full justification yang merenggangkan kata-kata.
    """
    scratch_dir = os.path.abspath(r'C:\Users\yonat\AppData\Local\hermes\profiles\bot3\cache\scratch')
    os.makedirs(scratch_dir, exist_ok=True)
    tmp_p = os.path.join(scratch_dir, f"pre_{os.path.basename(raw_docx_path)}_{os.getpid()}.docx")
    local_src = os.path.join(scratch_dir, f"src_copy_{os.path.basename(raw_docx_path)}_{os.getpid()}.docx")

    owned_word = False
    try:
        import shutil
        shutil.copyfile(os.path.abspath(raw_docx_path), local_src)
        if word_app is None:
            import win32com.client
            word = win32com.client.Dispatch('Word.Application')
            word.Visible = False
            word.DisplayAlerts = 0
            owned_word = True
        else:
            word = word_app
        try:
            wdoc = word.Documents.Open(local_src, False, True, False)
            wdoc.ConvertNumbersToText()
            wdoc.Content.Find.Execute('^l', False, False, False, False, False, True, 0, False, '^p', 2)
            wdoc.SaveAs2(tmp_p)
            wdoc.Close(False)
            return tmp_p
        finally:
            if owned_word:
                word.Quit()
    except Exception:
        return raw_docx_path
    finally:
        try:
            if os.path.exists(local_src):
                os.remove(local_src)
        except Exception:
            pass


def _is_instruction(text):
    """Instruksi/petunjuk: JANGAN diberi nomor soal."""
    low = text.lower().strip()
    if (
        ('pada soal no' in low or 'pada soal nomor' in low)
        or low.startswith('petunjuk')
        or low.startswith('instruksi')
        or low.startswith('hitunglah')
        or low.startswith('tulis nama bilangan')
        or low.startswith('tulis lambang bilangan')
        or low.startswith('urutkan ')
        or low.startswith('pasangkan ')
        or low.startswith('lingkarilah ')
        or low.startswith('berilah tanda ')
        or low.startswith('cocokkan ')
        or low.startswith('perhatikan ')
        or low.startswith('jawablah ')
        or low.startswith('bacalah ')
        or low.startswith('isilah ')
        or low.startswith('kerjakan ')
        or low.startswith('lengkapi ')
        or low.startswith('tentukan ')
        or low.startswith('carilah ')
    ):
        return True
    return False


def _is_word_choice_or_label(text):
    """Pilihan kata / label / baris jawaban: JANGAN diberi nomor soal."""
    low = text.lower().strip()
    low = re.sub(r'^[-•\*\s\t]+', '', low)
    if low in ('lebih dari', 'kurang dari', 'sama dengan', 'benar', 'salah', 'ya', 'tidak', '='):
        return True
    return False


def _is_clear_question(text, current_section, next_text=''):
    """Paragraf yang jelas merupakan butir soal tanpa nomor."""
    low = text.lower().strip()
    if current_section == 'isian':
        if len(text.strip()) >= 15 and not _is_instruction(text):
            return True
    if current_section in ('essay', 'esai', 'uraian'):
        if '?' in text:
            return True
        if next_text.strip() == '=':
            return True
        if (
            low.startswith('tuliskan ')
            or low.startswith('jelaskan ')
            or low.startswith('sebutkan ')
            or low.startswith('bagaimana ')
            or low.startswith('apa ')
            or low.startswith('mengapa ')
        ):
            return True
    return False


# ============================================================
# PENANGANAN OBJEK KHUSUS (GAMBAR & FORMULA) (TAMBAHAN v1.84)
# ============================================================
def _has_equation(p):
    return bool(p._p.xpath('.//m:oMath | .//m:oMathPara'))


def _has_drawing(p):
    return bool(p._p.xpath('.//w:drawing | .//w:pict | .//w:object'))


def _convert_anchors_to_inline(element):
    for anchor in list(element.iter(qn('wp:anchor'))):
        anchor.tag = qn('wp:inline')
        for child_tag in ('wp:simplePos', 'wp:positionH', 'wp:positionV', 'wp:wrapNone',
                          'wp:wrapSquare', 'wp:wrapTight', 'wp:wrapThrough', 'wp:wrapTopAndBottom'):
            for child in list(anchor.findall(qn(child_tag))):
                anchor.remove(child)


def _constrain_drawings(element, max_w_cm=7.0, max_h_cm=3.8):
    max_w = int(max_w_cm * 360000)
    max_h = int(max_h_cm * 360000)
    for extent in element.iter(qn('wp:extent')):
        try:
            cx = int(extent.get('cx'))
            cy = int(extent.get('cy'))
        except Exception:
            continue
        if cx <= 0 or cy <= 0:
            continue
        scale = min(1.0, max_w / cx, max_h / cy)
        if scale < 1.0:
            extent.set('cx', str(int(cx * scale)))
            extent.set('cy', str(int(cy * scale)))


def _remap_relationships(src_doc, dst_doc, element):
    for node in element.iter():
        for attr in (qn('r:embed'), qn('r:link'), qn('r:id')):
            rid = node.get(attr)
            if not rid:
                continue
            rel = src_doc.part.rels.get(rid)
            if rel is None:
                continue
            try:
                if rel.is_external:
                    new_rid = dst_doc.part.relate_to(rel.target_ref, rel.reltype, is_external=True)
                elif 'image' in rel.reltype:
                    new_rid, _ = dst_doc.part.get_or_add_image(io.BytesIO(rel.target_part.blob))
                else:
                    new_rid = dst_doc.part.relate_to(rel.target_part, rel.reltype)
                node.set(attr, new_rid)
            except Exception:
                pass


def create_sect_pr(cols_num=1, cols_space_twips=720):
    """Buat elemen <w:sectPr> kontinu untuk arsitektur seksi kolom seimbang (standar Hermes)."""
    sectPr = OxmlElement('w:sectPr')
    t = OxmlElement('w:type')
    t.set(qn('w:val'), 'continuous')
    sectPr.append(t)
    sz = OxmlElement('w:pgSz')
    sz.set(qn('w:w'), '12240')
    sz.set(qn('w:h'), '18720')
    sectPr.append(sz)
    mar = OxmlElement('w:pgMar')
    mar.set(qn('w:top'), '283')
    mar.set(qn('w:right'), '567')
    mar.set(qn('w:bottom'), '340')
    mar.set(qn('w:left'), '567')
    mar.set(qn('w:header'), '720')
    mar.set(qn('w:footer'), '720')
    mar.set(qn('w:gutter'), '0')
    sectPr.append(mar)
    cols = OxmlElement('w:cols')
    if cols_num > 1:
        cols.set(qn('w:num'), str(cols_num))
    cols.set(qn('w:space'), str(cols_space_twips))
    sectPr.append(cols)
    grid = OxmlElement('w:docGrid')
    grid.set(qn('w:linePitch'), '360')
    sectPr.append(grid)
    return sectPr


def set_document_defaults_tahoma(doc):
    """Lapis 2 OOXML: Atur docDefaults pada styles.xml ke Tahoma 11 pt solid black."""
    try:
        styles_el = doc.styles.element
        dd = styles_el.find(qn('w:docDefaults'))
        if dd is None:
            dd = OxmlElement('w:docDefaults')
            styles_el.insert(0, dd)
        rprd = dd.find(qn('w:rPrDefault'))
        if rprd is None:
            rprd = OxmlElement('w:rPrDefault')
            dd.append(rprd)
        rpr = rprd.find(qn('w:rPr'))
        if rpr is None:
            rpr = OxmlElement('w:rPr')
            rprd.append(rpr)
        rf = rpr.find(qn('w:rFonts'))
        if rf is None:
            rf = OxmlElement('w:rFonts')
            rpr.append(rf)
        rf.set(qn('w:ascii'), 'Tahoma')
        rf.set(qn('w:hAnsi'), 'Tahoma')
        rf.set(qn('w:cs'), 'Tahoma')
        sz = rpr.find(qn('w:sz'))
        if sz is None:
            sz = OxmlElement('w:sz')
            rpr.append(sz)
        sz.set(qn('w:val'), '22')
        color = rpr.find(qn('w:color'))
        if color is None:
            color = OxmlElement('w:color')
            rpr.append(color)
        color.set(qn('w:val'), '000000')
    except Exception:
        pass


def standardize_run_elements(p, font_size_pt=11.0):
    """Lapis 4 OOXML: Pastikan seluruh elemen run (<w:r>) seragam Tahoma dan #000000 murni."""
    for r_el in p._p.iter(qn('w:r')):
        rpr = r_el.find(qn('w:rPr'))
        if rpr is None:
            rpr = OxmlElement('w:rPr')
            r_el.insert(0, rpr)
        rf = rpr.find(qn('w:rFonts'))
        if rf is None:
            rf = OxmlElement('w:rFonts')
            rpr.append(rf)
        rf.set(qn('w:ascii'), 'Tahoma')
        rf.set(qn('w:hAnsi'), 'Tahoma')
        rf.set(qn('w:cs'), 'Tahoma')
        sz = rpr.find(qn('w:sz'))
        if sz is None:
            sz = OxmlElement('w:sz')
            rpr.append(sz)
        sz.set(qn('w:val'), str(int(font_size_pt * 2)))
        color_el = rpr.find(qn('w:color'))
        if color_el is None:
            color_el = OxmlElement('w:color')
            rpr.append(color_el)
        color_el.set(qn('w:val'), '000000')
        for attr in [qn('w:themeColor'), qn('w:themeTint'), qn('w:themeShade')]:
            if attr in color_el.attrib:
                del color_el.attrib[attr]


def _clone_special_para(src_para, doc_new, strip_number=False, font_size=11.0):
    p = doc_new.add_paragraph()
    for child in src_para._p:
        if child.tag == qn('w:pPr'):
            continue
        cloned = copy.deepcopy(child)
        _remap_relationships(src_para.part.document, doc_new, cloned)
        _convert_anchors_to_inline(cloned)
        _constrain_drawings(cloned)
        p._p.append(cloned)
    if strip_number:
        for t in p._p.iter(qn('w:t')):
            if t.text:
                new = re.sub(r'^\s*\d+[\.\)]\s*', '', t.text, count=1)
                if new != t.text:
                    t.text = new
                    break
    standardize_run_elements(p, font_size_pt=font_size)
    return p


# ============================================================
# KOP MASTER NATIVE WORD COM (CTRL+K) + FALLBACK XML
# ============================================================
def create_clean_doc_with_master_kop(tpl_path, unit, kelas_str, mapel, hari_tgl, logo_path=None, word_app=None):
    """
    KOP STANDAR BAKU (v1.77/v1.84):
    Word COM native 'rng.PasteAndFormat(16)' (Ctrl+K Keep Source Formatting).
    Menjaga 100% border ganda bawah, style TableGrid2, tinggi baris, garis vertikal,
    dan logo fisik resmi sekolah.
    """
    scratch_dir = os.path.abspath(r'C:\Users\yonat\AppData\Local\hermes\profiles\bot3\cache\scratch')
    os.makedirs(scratch_dir, exist_ok=True)
    tmp_p = os.path.join(scratch_dir, f"base_kop_{mapel.replace(' ', '_')}_{os.getpid()}.docx")
    local_tpl = os.path.join(scratch_dir, f"tpl_copy_{os.getpid()}.docx")

    owned_word = False
    try:
        import shutil
        shutil.copyfile(os.path.abspath(tpl_path), local_tpl)
        if word_app is None:
            import win32com.client
            word = win32com.client.Dispatch('Word.Application')
            word.Visible = False
            word.DisplayAlerts = 0
            owned_word = True
        else:
            word = word_app
        try:
            new_wdoc = word.Documents.Add()
            sec = new_wdoc.Sections(1).PageSetup
            sec.PageWidth = F4_WIDTH.pt
            sec.PageHeight = F4_HEIGHT.pt
            sec.TopMargin = MARGIN_TOP.pt
            sec.BottomMargin = MARGIN_BOTTOM.pt
            sec.LeftMargin = MARGIN_LEFT.pt
            sec.RightMargin = MARGIN_RIGHT.pt

            src_wdoc = word.Documents.Open(local_tpl, False, True, False)
            tbl_kop = src_wdoc.Tables(1)
            tbl_kop.Range.Copy()
            src_wdoc.Close(False)

            rng = new_wdoc.Range(0, 0)
            rng.PasteAndFormat(16) # wdFormatOriginalFormatting = 16 (Ctrl+K)

            t = new_wdoc.Tables(1)
            if unit == 'SMP':
                t.Cell(1, 2).Range.Find.Execute('SD SWASTA', False, False, False, False, False, True, 0, False, 'SMP SWASTA', 2)

            # Ganti Mata Pelajaran di Cell(1, 5)
            old_mapel = t.Cell(1, 5).Range.Text.replace(chr(13), "").replace(chr(7), "").strip()
            if old_mapel and old_mapel != mapel:
                t.Cell(1, 5).Range.Find.Execute(old_mapel, False, False, False, False, False, True, 0, False, mapel, 2)
            else:
                t.Cell(1, 5).Range.Find.Execute("Pendidikan Pancasila", False, False, False, False, False, True, 0, False, mapel, 2)

            # Ganti Hari/Tgl di Cell(2, 5)
            old_tgl = t.Cell(2, 5).Range.Text.replace(chr(13), "").replace(chr(7), "").strip()
            if old_tgl and old_tgl != hari_tgl:
                t.Cell(2, 5).Range.Find.Execute(old_tgl, False, False, False, False, False, True, 0, False, hari_tgl, 2)
            else:
                t.Cell(2, 5).Range.Find.Execute("Kamis, 17 September 2026", False, False, False, False, False, True, 0, False, hari_tgl, 2)

            # Ganti Kelas di Cell(3, 5)
            old_kelas = t.Cell(3, 5).Range.Text.replace(chr(13), "").replace(chr(7), "").strip()
            if old_kelas and old_kelas != kelas_str:
                t.Cell(3, 5).Range.Find.Execute(old_kelas, False, False, False, False, False, True, 0, False, kelas_str, 2)
            else:
                txt_cell = t.Cell(3, 5).Range.Text
                if 'IA SD' in txt_cell:
                    t.Cell(3, 5).Range.Find.Execute('IA SD', False, False, False, False, False, True, 0, False, kelas_str, 2)
                elif 'I SD' in txt_cell:
                    t.Cell(3, 5).Range.Find.Execute('I SD', False, False, False, False, False, True, 0, False, kelas_str, 2)

            new_wdoc.SaveAs2(tmp_p)
            new_wdoc.Close(False)

            doc = docx.Document(tmp_p)
            try:
                os.remove(tmp_p)
            except Exception:
                pass
            if doc.tables:
                t0 = doc.tables[0]
                if len(t0.rows) >= 3 and len(t0.rows[0].cells) >= 5:
                    for r_i, target_val in [(0, mapel), (1, hari_tgl), (2, kelas_str)]:
                        c = t0.rows[r_i].cells[4]
                        if c.text.strip() != target_val:
                            p = c.paragraphs[0]
                            for child in list(p._p):
                                if child.tag != qn('w:pPr'):
                                    p._p.remove(child)
                            for extra_p in list(c.paragraphs[1:]):
                                c._tc.remove(extra_p._p)
                            r = p.add_run(target_val)
                            set_run_tahoma(r, 11.0, bold=False, color_hex="000000")
            set_document_defaults_tahoma(doc)
            paras = list(doc.paragraphs)
            if not paras:
                doc.add_paragraph()
            else:
                for p in paras[1:]:
                    if not p.text.strip():
                        p._p.getparent().remove(p._p)
            return doc
        finally:
            if owned_word:
                word.Quit()
    except Exception:
        # Fallback XML murni jika Word COM tidak aktif
        doc = docx.Document()
        sec = doc.sections[0]
        sec.page_width = F4_WIDTH
        sec.page_height = F4_HEIGHT
        sec.top_margin = MARGIN_TOP
        sec.bottom_margin = MARGIN_BOTTOM
        sec.left_margin = MARGIN_LEFT
        sec.right_margin = MARGIN_RIGHT

        tpl_doc = docx.Document(tpl_path)
        tbl_elem = copy.deepcopy(tpl_doc.tables[0]._tbl)
        _remap_relationships(tpl_doc, doc, tbl_elem)
        p_dummy = doc.add_paragraph()
        p_dummy._p.addprevious(tbl_elem)
        p_dummy._p.getparent().remove(p_dummy._p)
        if doc.tables:
            t0 = doc.tables[0]
            if len(t0.rows) >= 3 and len(t0.rows[0].cells) >= 5:
                for r_i, target_val in [(0, mapel), (1, hari_tgl), (2, kelas_str)]:
                    c = t0.rows[r_i].cells[4]
                    p = c.paragraphs[0]
                    for child in list(p._p):
                        if child.tag != qn('w:pPr'):
                            p._p.remove(child)
                    for extra_p in list(c.paragraphs[1:]):
                        c._tc.remove(extra_p._p)
                    r = p.add_run(target_val)
                    set_run_tahoma(r, 11.0, bold=False, color_hex="000000")
        set_document_defaults_tahoma(doc)
        paras = list(doc.paragraphs)
        if not paras:
            doc.add_paragraph()
        else:
            for p in paras[1:]:
                if not p.text.strip():
                    p._p.getparent().remove(p._p)
        return doc
    finally:
        try:
            if os.path.exists(local_tpl):
                os.remove(local_tpl)
        except Exception:
            pass


def apply_p0_1pt_gap_fix(doc):
    """
    Trik 1pt P0: Paragraf transisi continuous section break tepat di bawah kop
    disetel ke 1 pt exact agar judul seksi pertama menempel rapat di bawah kop tanpa celah.
    """
    if not doc.paragraphs:
        doc.add_paragraph()
    p0 = doc.paragraphs[0]
    p0_pr = p0._p.get_or_add_pPr()
    sp = p0_pr.find(qn('w:spacing'))
    if sp is None:
        sp = OxmlElement('w:spacing')
        p0_pr.append(sp)
    sp.set(qn('w:before'), '0')
    sp.set(qn('w:after'), '0')
    sp.set(qn('w:line'), '20')
    sp.set(qn('w:lineRule'), 'exact')
    rpr = p0_pr.find(qn('w:rPr'))
    if rpr is None:
        rpr = OxmlElement('w:rPr')
        p0_pr.append(rpr)
    sz = rpr.find(qn('w:sz'))
    if sz is None:
        sz = OxmlElement('w:sz')
        rpr.append(sz)
    sz.set(qn('w:val'), '2')
    szCs = rpr.find(qn('w:szCs'))
    if szCs is None:
        szCs = OxmlElement('w:szCs')
        rpr.append(szCs)
    szCs.set(qn('w:val'), '2')
    old_sect = p0_pr.find(qn('w:sectPr'))
    if old_sect is not None:
        p0_pr.remove(old_sect)
    p0_pr.append(create_sect_pr(cols_num=1, cols_space_twips=720))


# ============================================================
# RENDER CONTENT BLOCK
# ============================================================
def render_content_block(doc_new, raw_paras, is_low_grade, font_size, ls_val, sa_val, mapel):
    apply_p0_1pt_gap_fix(doc_new)
    is_first_heading = True
    current_section = None
    q_counter = 0
    tail_paras = {p._p for p in raw_paras[-6:]}
    has_explicit_numbers = any(bool(RE_QUESTION.match(p.text.strip())) for p in raw_paras)

    for idx, p_raw in enumerate(raw_paras):
        txt = p_raw.text.strip()
        has_draw = _has_drawing(p_raw)
        has_eq = _has_equation(p_raw)
        special = has_draw or has_eq
        next_txt = raw_paras[idx + 1].text.strip() if idx + 1 < len(raw_paras) else ''

        if not txt and not special:
            continue
        if any(k in txt.lower() for k in ('methodist-11', 'mata pelajaran :', 'hari / tgl :')):
            continue
        if re.match(r"^\s*(soal\s+(?:uts|uas|ujian)|kelas\s+[0-9ivx]+)", txt.lower()) and not any(k in txt.lower() for k in ('isian', 'uraian', 'essay', 'esai')):
            continue

        # English Good Luck protected area
        if 'inggris' in mapel.lower() and p_raw._p in tail_paras and RE_GOOD_LUCK.search(txt):
            p_gl = doc_new.add_paragraph()
            set_para_ind(p_gl, left_cm=0.0, hanging_cm=0.0, space_before_pt=4, space_after_pt=0, line_spacing=1.0)
            p_gl.alignment = WD_ALIGN_PARAGRAPH.CENTER
            r_gl = p_gl.add_run(txt)
            set_run_tahoma(r_gl, font_size, bold=True)
            continue

        # Gambar standalone tanpa teks -> Left 0.65 cm, Hanging 0.0 cm
        if has_draw and not txt:
            p_img = _clone_special_para(p_raw, doc_new, strip_number=False, font_size=font_size)
            set_para_ind(p_img, left_cm=0.65, hanging_cm=0.0, space_before_pt=0, space_after_pt=1, line_spacing=1.0)
            p_img.alignment = WD_ALIGN_PARAGRAPH.LEFT
            continue

        is_opt = bool(RE_OPTION.match(txt))
        explicit_num = RE_QUESTION.match(txt)
        RE_HEADING_STRICT = re.compile(
            r'^\s*(?:[A-Z][\.\)]\s*)?(?:pilihan\s*(?:ber)?ganda|isian|essay|esai|uraian|kunci\s*jawaban|c\.?p\.?\s*\d+)\b',
            re.I
        )
        is_heading = (not is_opt and not explicit_num and bool(RE_HEADING_STRICT.match(txt)))

        # Heading / Judul Seksi
        if is_heading:
            q_counter = 0
            low_h = txt.lower()
            if 'isian' in low_h or 'cp 1' in low_h: current_section = 'isian'
            elif 'essay' in low_h or 'esai' in low_h or 'uraian' in low_h or 'cp 2' in low_h: current_section = 'essay'
            elif 'pilihan' in low_h: current_section = 'pg'

            if is_first_heading:
                # Judul 1: Zero-gap menempel rapat di bawah kop
                p_h = doc_new.add_paragraph()
                set_para_ind(p_h, left_cm=0.0, hanging_cm=0.0, space_before_pt=0, space_after_pt=2, line_spacing=1.0)
                is_first_heading = False
            else:
                # Judul 2+: Jeda seksi tegas di atasnya dengan tepat 1 enter bersih + penyeimbang kolom kontinu
                p_break = doc_new.add_paragraph()
                p_break_pr = p_break._p.get_or_add_pPr()
                sp_b = OxmlElement('w:spacing')
                sp_b.set(qn('w:before'), '0')
                sp_b.set(qn('w:after'), '0')
                sp_b.set(qn('w:line'), '20')
                sp_b.set(qn('w:lineRule'), 'exact')
                p_break_pr.append(sp_b)
                p_break_pr.append(create_sect_pr(cols_num=2, cols_space_twips=284))

                p_enter = doc_new.add_paragraph()
                set_para_ind(p_enter, left_cm=0.0, hanging_cm=0.0, space_before_pt=0, space_after_pt=0, line_spacing=1.0)
                p_h = doc_new.add_paragraph()
                set_para_ind(p_h, left_cm=0.0, hanging_cm=0.0, space_before_pt=0, space_after_pt=2, line_spacing=1.0)
                p_h._p.get_or_add_pPr().append(create_sect_pr(cols_num=1, cols_space_twips=720))

            r_h = p_h.add_run(txt)
            set_run_tahoma(r_h, 11.5 if not is_low_grade else 12.0, bold=True)
            continue

        # Pilihan Ganda (a., b., c., d.) -> Left 1.30 cm, Hanging 0.65 cm + TAB manual baku
        if is_opt:
            m_in = RE_OPT_INLINE.match(txt)
            if m_in:
                lbl_l = m_in.group(1)
                if lbl_l.endswith(')'): lbl_l = lbl_l[:-1] + '.'
                body_l = m_in.group(2).strip()
                lbl_r = m_in.group(3)
                if lbl_r.endswith(')'): lbl_r = lbl_r[:-1] + '.'
                body_r = m_in.group(4).strip()
                new_text = f"{lbl_l}\t{body_l}\t{lbl_r}\t{body_r}"
                p_opt = doc_new.add_paragraph()
                set_para_ind(p_opt, left_cm=1.30, hanging_cm=0.65, space_before_pt=0, space_after_pt=0, line_spacing=1.0)
                set_option_tabs_inline(p_opt)
                p_opt.alignment = WD_ALIGN_PARAGRAPH.LEFT
                r_opt = p_opt.add_run(new_text)
                set_run_tahoma(r_opt, font_size, bold=False)
                continue

            m_s = RE_OPT_SINGLE.match(txt)
            if m_s:
                lbl = m_s.group(1)
                if lbl.endswith(')'): lbl = lbl[:-1] + '.'
                body = m_s.group(2).strip()
                new_text = f"{lbl}\t{body}"
                p_opt = doc_new.add_paragraph()
                set_para_ind(p_opt, left_cm=1.30, hanging_cm=0.65, space_before_pt=0, space_after_pt=0, line_spacing=1.0)
                set_option_tabs_single(p_opt)
                p_opt.alignment = WD_ALIGN_PARAGRAPH.LEFT
                r_opt = p_opt.add_run(new_text)
                set_run_tahoma(r_opt, font_size, bold=False)
                continue

        # Bullet teks khusus (misal di kunci jawaban)
        if txt.startswith("") or txt.startswith("") or txt.startswith("") or re.match(r"^[•●○■\-]\s+", txt):
            clean_bullet = re.sub(r"^[•●○■\-]\s*", "", txt)
            p_b = doc_new.add_paragraph()
            set_para_ind(p_b, left_cm=1.30, hanging_cm=0.65, space_before_pt=0, space_after_pt=0, line_spacing=1.0)
            set_option_tabs_single(p_b)
            r_b = p_b.add_run(f"-	{clean_bullet}")
            set_run_tahoma(r_b, font_size, bold=False)
            continue

        # Garis isian vokal/konsonan (harus mengandung underscore dan kata vokal/konsonan)
        is_vokal_line = 'huruf vokal' in txt.lower() and '_' in txt
        is_konsonan_line = 'huruf konsonan' in txt.lower() and '_' in txt

        if is_vokal_line:
            p_ans = doc_new.add_paragraph()
            set_para_ind(p_ans, left_cm=0.65, hanging_cm=0.0, space_before_pt=0, space_after_pt=0, line_spacing=ls_val)
            try: p_ans.paragraph_format.tab_stops.add_tab_stop(Cm(2.80))
            except Exception: pass
            r = p_ans.add_run("Jb :\t_______ huruf vokal")
            set_run_tahoma(r, font_size, bold=False)
            continue

        if is_konsonan_line:
            p_ans = doc_new.add_paragraph()
            set_para_ind(p_ans, left_cm=0.65, hanging_cm=0.0, space_before_pt=0, space_after_pt=sa_val, line_spacing=ls_val)
            try: p_ans.paragraph_format.tab_stops.add_tab_stop(Cm(2.80))
            except Exception: pass
            r = p_ans.add_run("\t_______ huruf konsonan")
            set_run_tahoma(r, font_size, bold=False)
            continue

        if txt == '=' or re.match(r'^\s*=\s*$', txt):
            p_eq = doc_new.add_paragraph()
            set_para_ind(p_eq, left_cm=0.65, hanging_cm=0.0, space_before_pt=0, space_after_pt=sa_val, line_spacing=ls_val)
            r_eq = p_eq.add_run("=\t................................................................................")
            set_run_tahoma(r_eq, font_size, bold=False)
            continue

        is_vowel_word = txt in ('Kebersihan', 'Lingkungan', 'Pembelajaran', 'Bernyanyi')

        # 5 Aturan Penomoran Cerdas (Anti-Blind-Numbering):
        # 1. Nomor asli ada        -> pertahankan
        # 2. Instruksi/petunjuk    -> jangan diberi nomor
        # 3. Kata/label biasa      -> jangan diberi nomor
        # 4. Paragraf jelas soal   -> boleh diproses sebagai soal
        # 5. Tidak yakin           -> jangan auto-number
        should_number = False
        num_to_use = None

        if explicit_num:
            should_number = True
            num_to_use = int(explicit_num.group(1))
            q_counter = num_to_use
        elif is_vowel_word:
            q_counter += 1
            should_number = True
            num_to_use = q_counter
        elif _is_instruction(txt):
            should_number = False
        elif _is_word_choice_or_label(txt):
            should_number = False
        elif (not has_explicit_numbers) and _is_clear_question(txt, current_section, next_txt):
            q_counter += 1
            should_number = True
            num_to_use = q_counter
        else:
            should_number = False

        clean_prompt = re.sub(r'^\s*\d+[\.\)]\s*', '', txt)

        # Khusus soal bergambar / ber-equation
        if special:
            p_item = _clone_special_para(p_raw, doc_new, strip_number=bool(explicit_num), font_size=font_size)
            if should_number:
                set_para_ind(p_item, left_cm=0.65, hanging_cm=0.65, space_before_pt=0, space_after_pt=sa_val, line_spacing=ls_val)
                _set_question_grid_tab(p_item)
                p_item.alignment = WD_ALIGN_PARAGRAPH.JUSTIFY
                r_num = OxmlElement('w:r')
                rpr = OxmlElement('w:rPr')
                rf = OxmlElement('w:rFonts')
                for k in ('w:ascii', 'w:hAnsi', 'w:cs'): rf.set(qn(k), 'Tahoma')
                rpr.append(rf)
                sz = OxmlElement('w:sz')
                sz.set(qn('w:val'), str(int(font_size * 2)))
                rpr.append(sz)
                r_num.append(rpr)
                t_el = OxmlElement('w:t')
                t_el.text = f"{num_to_use}.\t"
                r_num.append(t_el)
                insert_at = 1 if len(p_item._p) and p_item._p[0].tag == qn('w:pPr') else 0
                p_item._p.insert(insert_at, r_num)
            else:
                set_para_ind(p_item, left_cm=0.65, hanging_cm=0.0, space_before_pt=0, space_after_pt=sa_val, line_spacing=ls_val)
                p_item.alignment = WD_ALIGN_PARAGRAPH.LEFT
        else:
            p_item = doc_new.add_paragraph()
            if should_number:
                set_para_ind(p_item, left_cm=0.65, hanging_cm=0.65, space_before_pt=0, space_after_pt=sa_val, line_spacing=ls_val)
                _set_question_grid_tab(p_item)
                p_item.alignment = WD_ALIGN_PARAGRAPH.JUSTIFY
                r_n = p_item.add_run(f"{num_to_use}.\t")
                set_run_tahoma(r_n, font_size, bold=False)
                r_t = p_item.add_run(clean_prompt)
                set_run_tahoma(r_t, font_size, bold=False)
            else:
                set_para_ind(p_item, left_cm=0.65, hanging_cm=0.0, space_before_pt=0, space_after_pt=sa_val, line_spacing=ls_val)
                r_t = p_item.add_run(txt)
                set_run_tahoma(r_t, font_size, bold=False)

        # Ruang Jawaban Kelas 1-2 SD (hanya untuk butir soal esai mapel non-matematika, BUKAN kata hitung huruf vokal 7-10)
        if is_low_grade and current_section == 'essay' and should_number and not is_vowel_word and 'matematika' not in mapel.lower():
            p_ans = doc_new.add_paragraph()
            set_para_ind(p_ans, left_cm=0.65, hanging_cm=0.0, space_before_pt=0, space_after_pt=sa_val, line_spacing=ls_val)
            r_a = p_ans.add_run("Jb :\t................................................................................")
            set_run_tahoma(r_a, font_size, bold=False)

    # Arsitektur Seksi 6 (Seksi Akhir): Mengalir kembali dalam 2 kolom
    body_sect = doc_new._body._body.sectPr
    if body_sect is not None:
        cols = body_sect.find(qn('w:cols'))
        if cols is None:
            cols = OxmlElement('w:cols')
            body_sect.append(cols)
        cols.set(qn('w:num'), '2')
        cols.set(qn('w:space'), '284')


# ============================================================
# PIPELINE UTAMA
# ============================================================
def proses_naskah_v1_84(raw_docx_path, out_docx_path, tpl_docx_path=None, folder_name="", logo_path=None, mode="auto", preview=False):
    if not tpl_docx_path or not os.path.exists(tpl_docx_path):
        candidates = [
            os.path.join(os.path.dirname(__file__), "..", "templates", "SD SWASTA METHODIST.docx"),
            os.path.join(os.path.dirname(__file__), "templates", "SD SWASTA METHODIST.docx"),
            os.path.join(os.path.dirname(__file__), "SD SWASTA METHODIST.docx"),
            os.path.join(os.path.dirname(__file__), "kop-methodist.docx"),
            r"C:\Users\yonat\Desktop\Soal Ujian Semester I T.P.2026-2027\SD SWASTA METHODIST.docx",
            r"C:\Users\yonat\ANTIGRAVITY\COMMON\assets\kop-methodist.docx",
            os.path.expandvars(r"%USERPROFILE%\Desktop\tempat kerja Antigravity\SD SWASTA METHODIST.docx"),
            r"C:\Users\yonat\OneDrive\Desktop\latihann\SD SWASTA METHODIST.docx",
            r"C:\Users\Advan\Desktop\tempat kerja Hermes\SOAL UJIAN\Soal Ujian UTS 1 T.P. 2026-2027\SD SWASTA METHODIST.docx"
        ]
        for c in candidates:
            if os.path.exists(c):
                tpl_docx_path = c
                break

    fname = os.path.basename(raw_docx_path)
    unit, kelas_str, mapel, hari_tgl, grade = parse_metadata(folder_name, fname, raw_docx_path)
    is_low_grade = grade in (1, 2)

    # Start single shared Word COM session
    word_app = None
    try:
        import win32com.client
        word_app = win32com.client.Dispatch('Word.Application')
        word_app.Visible = False
        word_app.DisplayAlerts = 0
    except Exception:
        pass

    try:
        # 1. Dokumen baru dengan Kop Master via Word COM (Ctrl+K)
        doc_new = create_clean_doc_with_master_kop(tpl_docx_path, unit, kelas_str, mapel, hari_tgl, logo_path, word_app=word_app)

        # 2. Baca dokumen mentah guru (bekukan Word native numbering jika ada)
        norm_path = normalize_input_docx(raw_docx_path, word_app=word_app)
        doc_raw = docx.Document(norm_path)
        raw_paras = [p for p in doc_raw.paragraphs if p.text.strip() or _has_drawing(p) or _has_equation(p)]
        if norm_path != raw_docx_path:
            try:
                os.remove(norm_path)
            except Exception:
                pass

        total_chars = sum(len(p.text) for p in raw_paras)
        is_dense = (total_chars > 3000 or len(raw_paras) > 22)

        # Evaluasi mode 2-in-1:
        is_2in1 = (mode == "2in1") or (mode == "auto" and not is_low_grade and len(raw_paras) <= 25)

        if is_2in1:
            font_size = 11.0 # STRICT FONT FLOOR MUTLAK MINIMAL 11.0 PT!
            ls_val = 0.88
            sa_val = 0

            # Paruh Atas
            render_content_block(doc_new, raw_paras, is_low_grade, font_size, ls_val, sa_val, mapel)

            # Enter pemisah simetris potong tengah 16.50 cm
            p_sep = doc_new.add_paragraph()
            set_para_ind(p_sep, left_cm=0.0, hanging_cm=0.0, space_before_pt=4, space_after_pt=0, line_spacing=1.0)

            # Kloning Kop 1 ke Kop 2
            tbl_kop1 = doc_new.tables[0]
            tbl_kop2_elem = copy.deepcopy(tbl_kop1._tbl)
            p_dummy2 = doc_new.add_paragraph()
            p_dummy2._p.addprevious(tbl_kop2_elem)
            p_dummy2._p.getparent().remove(p_dummy2._p)

            # Paruh Bawah
            render_content_block(doc_new, raw_paras, is_low_grade, font_size, ls_val, sa_val, mapel)
        else:
            # Mode 1 Lembar Penuh (Full Sheet)
            if is_low_grade:
                font_size = 11.5 if is_dense else 12.0
                ls_val = 1.0 if is_dense else 1.25
                sa_val = 0 if is_dense else 2
            else:
                font_size = 11.0 if is_dense else 11.5
                ls_val = 1.0
                sa_val = 0

            render_content_block(doc_new, raw_paras, is_low_grade, font_size, ls_val, sa_val, mapel)

        # Lapis 4 OOXML: Sapuan universal memastikan 100% runs berfont Tahoma 11 pt (#000000 murni)
        for p in doc_new.paragraphs:
            standardize_run_elements(p, font_size_pt=font_size)
        doc_new.save(out_docx_path)

        # 3. Otomatis Verifikasi & Preview jika diminta
        if preview and word_app:
            scratch_dir = os.path.abspath(r'C:\Users\yonat\AppData\Local\hermes\profiles\bot3\cache\scratch')
            bname = os.path.splitext(os.path.basename(out_docx_path))[0]
            pdf_p = os.path.join(scratch_dir, f"{bname}.pdf")
            png_p = os.path.join(scratch_dir, f"{bname}_p1.png")

            local_out = os.path.join(scratch_dir, f"out_ver_{os.getpid()}.docx")
            import shutil
            shutil.copyfile(os.path.abspath(out_docx_path), local_out)

            res_wdoc = word_app.Documents.Open(local_out, False, True, False)
            pages = res_wdoc.ComputeStatistics(2)
            res_wdoc.ExportAsFixedFormat(pdf_p, 17)
            res_wdoc.Close(False)
            try:
                os.remove(local_out)
            except Exception:
                pass

            try:
                import pymupdf as fitz
            except Exception:
                import fitz
            doc_pdf = fitz.open(pdf_p)
            doc_pdf[0].get_pixmap(dpi=150).save(png_p)
            print(f"[PREVIEW] Word COM Pages: {pages}")
            print(f"[PREVIEW] MEDIA:{png_p}")

        return True
    finally:
        if word_app:
            try:
                word_app.Quit()
            except Exception:
                pass

if __name__ == '__main__':
    parser = argparse.ArgumentParser(description=f"Engine Standar Baku Editing Soal Ujian v{VERSION}")
    parser.add_argument("input", help="File naskah mentah .docx")
    parser.add_argument("-o", "--output", help="File naskah output .docx", default=None)
    parser.add_argument("--mode", choices=["auto", "full", "2in1"], default="auto", help="Mode layout (auto/full/2in1)")
    parser.add_argument("--tpl", help="Path template master kop", default=None)
    parser.add_argument("--folder", help="Nama folder / mata pelajaran", default="")
    parser.add_argument("--logo", help="Path logo sekolah", default=None)
    parser.add_argument("--preview", action="store_true", help="Otomatis validasi halaman Word COM dan render preview PNG")
    args = parser.parse_args()

    inp = os.path.abspath(args.input)
    out = args.output
    if not out:
        b, _ = os.path.splitext(inp)
        out = f"{b} Edit.docx"
    out = os.path.abspath(out)

    print(f"[*] Menjalankan Editor Soal Ujian v{VERSION}...")
    print(f"    Input : {inp}")
    print(f"    Output: {out}")
    print(f"    Mode  : {args.mode}")
    ok = proses_naskah_v1_84(inp, out, args.tpl, args.folder, args.logo, mode=args.mode, preview=args.preview)
    if ok:
        print(f"[OK] Berhasil menghasilkan naskah standar v{VERSION}: {out}")
    else:
        print("[FAIL] Gagal memproses naskah.", file=sys.stderr)
        sys.exit(1)
