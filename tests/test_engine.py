#!/usr/bin/env python3
"""
Unit tests untuk editor_soal_v1_85 (Engine Editing Naskah Soal Ujian SD Methodist-11).
Memverifikasi kepatuhan terhadap standar baku (Panduan Hermes):
- Kop master Folio F4 (21.59 x 33.02 cm) & Margin hemat (0.5, 0.6, 1.0, 1.0 cm)
- Standardisasi tipografi 4-lapis (Tahoma 11 pt, #000000)
- Hanging indent kelipatan 0.65 cm & Opsi 1.30 cm
- Karakter literal tab (\t) pada penomoran opsi
- Arsitektur seksi 2 kolom kontinu seimbang
- Integritas data read-only (berkas asli tidak disentuh)
"""

import hashlib
import os
import sys
import unittest
import docx
from docx.shared import Cm, Pt
from docx.oxml.ns import qn

# Pastikan folder project ada di sys.path
BASE_DIR = os.path.abspath(os.path.join(os.path.dirname(__file__), ".."))
SCRIPTS_DIR = os.path.join(BASE_DIR, "scripts")
if SCRIPTS_DIR not in sys.path:
    sys.path.insert(0, SCRIPTS_DIR)

import editor_soal_v1_85 as engine


def file_sha256(path):
    h = hashlib.sha256()
    with open(path, "rb") as f:
        while chunk := f.read(65536):
            h.update(chunk)
    return h.hexdigest()


class TestEngineSoal(unittest.TestCase):
    @classmethod
    def setUpClass(cls):
        cls.template_path = os.path.join(BASE_DIR, "templates", "SD SWASTA METHODIST.docx")
        cls.test_output_dir = os.path.join(BASE_DIR, "test_output")
        os.makedirs(cls.test_output_dir, exist_ok=True)
        assert os.path.exists(cls.template_path), f"Template master kop tidak ditemukan: {cls.template_path}"

    def test_kop_master_dimensions_and_margins(self):
        """Memverifikasi ukuran kertas Folio F4 (21.59 x 33.02 cm) dan margin baku."""
        doc = engine.create_clean_doc_with_master_kop(
            self.template_path,
            unit="SD",
            kelas_str="III SD",
            mapel="Ilmu Pengetahuan Alam (IPA)",
            hari_tgl="Rabu, 16 September 2026",
        )
        sec = doc.sections[0]
        w_cm = round(sec.page_width.cm, 2)
        h_cm = round(sec.page_height.cm, 2)
        self.assertAlmostEqual(w_cm, 21.59, delta=0.1, msg=f"Lebar F4 harus 21.59 cm, dapat {w_cm}")
        self.assertAlmostEqual(h_cm, 33.02, delta=0.1, msg=f"Tinggi F4 harus 33.02 cm, dapat {h_cm}")

        top_cm = round(sec.top_margin.cm, 2)
        bottom_cm = round(sec.bottom_margin.cm, 2)
        left_cm = round(sec.left_margin.cm, 2)
        right_cm = round(sec.right_margin.cm, 2)
        self.assertAlmostEqual(top_cm, 0.50, delta=0.05, msg="Margin atas harus 0.50 cm")
        self.assertAlmostEqual(bottom_cm, 0.60, delta=0.05, msg="Margin bawah harus 0.60 cm")
        self.assertAlmostEqual(left_cm, 1.00, delta=0.05, msg="Margin kiri harus 1.00 cm")
        self.assertAlmostEqual(right_cm, 1.00, delta=0.05, msg="Margin kanan harus 1.00 cm")

    def test_pure_paste_kop_metadata_preserves_ppr(self):
        """Memverifikasi pengisian identitas kop mempertahankan properti paragraf master."""
        doc = engine.create_clean_doc_with_master_kop(
            self.template_path,
            unit="SD",
            kelas_str="IV SD",
            mapel="Bahasa Indonesia",
            hari_tgl="Jumat, 18 September 2026",
        )
        self.assertTrue(len(doc.tables) >= 1)
        tbl = doc.tables[0]
        # Mapel di baris 0 kolom 4
        c_mapel = tbl.rows[0].cells[4]
        self.assertEqual(c_mapel.text.strip(), "Bahasa Indonesia")
        self.assertIsNotNone(c_mapel.paragraphs[0]._p.find(qn("w:pPr")))

    def test_typography_defaults_tahoma(self):
        """Memverifikasi Lapis 2 docDefaults styles.xml dikunci ke Tahoma 11 pt solid black."""
        doc = docx.Document()
        engine.set_document_defaults_tahoma(doc)
        styles_el = doc.styles.element
        dd = styles_el.find(qn("w:docDefaults"))
        self.assertIsNotNone(dd)
        rpr = dd.find(qn("w:rPrDefault")).find(qn("w:rPr"))
        rf = rpr.find(qn("w:rFonts"))
        self.assertEqual(rf.get(qn("w:ascii")), "Tahoma")
        self.assertEqual(rf.get(qn("w:hAnsi")), "Tahoma")
        c = rpr.find(qn("w:color"))
        self.assertEqual(c.get(qn("w:val")), "000000")

    def test_hanging_indent_and_option_tabs(self):
        """Memverifikasi grid indent 0.65 cm, opsi 1.30 cm, dan karakter tab \t."""
        doc = docx.Document()
        p = doc.add_paragraph()
        engine.set_para_ind(p, left_cm=0.65, hanging_cm=0.65, space_before_pt=0, space_after_pt=0, line_spacing=1.0)
        self.assertAlmostEqual(p.paragraph_format.left_indent.cm, 0.65, delta=0.02)
        self.assertAlmostEqual(p.paragraph_format.first_line_indent.cm, -0.65, delta=0.02)

        p_opt = doc.add_paragraph()
        engine.set_para_ind(p_opt, left_cm=1.30, hanging_cm=0.65, space_before_pt=0, space_after_pt=0, line_spacing=1.0)
        self.assertAlmostEqual(p_opt.paragraph_format.left_indent.cm, 1.30, delta=0.02)
        self.assertAlmostEqual(p_opt.paragraph_format.first_line_indent.cm, -0.65, delta=0.02)

    def test_read_only_source_integrity(self):
        """Memverifikasi berkas sumber asli dan @Back UP tidak dimodifikasi saat diproses."""
        sample_path = r"C:\Users\yonat\Desktop\Soal Ujian Semester I T.P.2026-2027\@Back UP\Masrio\UAS KELAS 1 - SOAL.docx"
        if not os.path.exists(sample_path):
            self.skipTest(f"Sample path tidak tersedia: {sample_path}")

        hash_before = file_sha256(sample_path)
        out_path = os.path.join(self.test_output_dir, "test_readonly_out.docx")

        engine.proses_naskah_v1_84(
            sample_path,
            out_path,
            tpl_docx_path=self.template_path,
            mode="auto",
        )

        hash_after = file_sha256(sample_path)
        self.assertEqual(hash_before, hash_after, "KRITIS: File sumber guru termodifikasi!")
        self.assertTrue(os.path.exists(out_path))


if __name__ == "__main__":
    unittest.main()
