#!/usr/bin/env python3
"""
Test audit otomatis terhadap hasil engine pada sampel Kelas 1-6 SD.
Memverifikasi tercapainya target 0 FAIL, 0 WARN (100% lolos audit).
"""

import os
import sys
import unittest

BASE_DIR = os.path.abspath(os.path.join(os.path.dirname(__file__), ".."))
SCRIPTS_DIR = os.path.join(BASE_DIR, "scripts")
if SCRIPTS_DIR not in sys.path:
    sys.path.insert(0, SCRIPTS_DIR)

import editor_soal_v1_85 as engine
import audit_naskah as auditor

SAMPLES_MAP = {
    1: r"C:\Users\yonat\Desktop\Soal Ujian Semester I T.P.2026-2027\@Back UP\Masrio\UAS KELAS 1 - SOAL.docx",
    2: r"C:\Users\yonat\Desktop\Soal Ujian Semester I T.P.2026-2027\@Back UP\Masrio\UAS KELAS 2 - SOAL.docx",
    3: r"C:\Users\yonat\Desktop\Soal Ujian Semester I T.P.2026-2027\@Back UP\Jellys\IPA 3 Jellys 2026 S1 - soal.docx",
    4: r"C:\Users\yonat\Desktop\Soal Ujian Semester I T.P.2026-2027\@Back UP\Chelsea\US 1 B. Indo Kelas IV (blm edit tgl).docx",
    5: r"C:\Users\yonat\Desktop\Soal Ujian Semester I T.P.2026-2027\@Back UP\Donny\Soal UAS 1 Mate 5 sem 1.docx",
    6: r"C:\Users\yonat\Desktop\Soal Ujian Semester I T.P.2026-2027\@Back UP\Steven\UAS Matematika (6) T.P. 2026-2027.docx",
}


class TestAuditKelas1Sampai6(unittest.TestCase):
    @classmethod
    def setUpClass(cls):
        cls.template_path = os.path.join(BASE_DIR, "templates", "SD SWASTA METHODIST.docx")
        cls.output_dir = os.path.join(BASE_DIR, "test_output")
        os.makedirs(cls.output_dir, exist_ok=True)

    def _run_and_audit(self, grade, sample_path):
        if not os.path.exists(sample_path):
            self.skipTest(f"Sample kelas {grade} tidak ditemukan: {sample_path}")
        out_path = os.path.join(self.output_dir, f"audit_kelas_{grade}.docx")
        ok = engine.proses_naskah_v1_84(
            sample_path,
            out_path,
            tpl_docx_path=self.template_path,
            mode="auto",
        )
        self.assertTrue(ok, f"Engine gagal memproses naskah kelas {grade}")
        res = auditor.audit(out_path, std="SD")
        self.assertEqual(res["fail"], 0, f"Kelas {grade} menghasilkan {res['fail']} FAIL: {res['notes']}")
        self.assertEqual(res["warn"], 0, f"Kelas {grade} menghasilkan {res['warn']} WARN: {res['notes']}")

    def test_audit_kelas_1(self):
        """Audit naskah Kelas 1 SD: Target 0 FAIL, 0 WARN."""
        self._run_and_audit(1, SAMPLES_MAP[1])

    def test_audit_kelas_2(self):
        """Audit naskah Kelas 2 SD: Target 0 FAIL, 0 WARN."""
        self._run_and_audit(2, SAMPLES_MAP[2])

    def test_audit_kelas_3(self):
        """Audit naskah Kelas 3 SD: Target 0 FAIL, 0 WARN."""
        self._run_and_audit(3, SAMPLES_MAP[3])

    def test_audit_kelas_4(self):
        """Audit naskah Kelas 4 SD: Target 0 FAIL, 0 WARN."""
        self._run_and_audit(4, SAMPLES_MAP[4])

    def test_audit_kelas_5(self):
        """Audit naskah Kelas 5 SD: Target 0 FAIL, 0 WARN."""
        self._run_and_audit(5, SAMPLES_MAP[5])

    def test_audit_kelas_6(self):
        """Audit naskah Kelas 6 SD: Target 0 FAIL, 0 WARN."""
        self._run_and_audit(6, SAMPLES_MAP[6])


if __name__ == "__main__":
    unittest.main()
