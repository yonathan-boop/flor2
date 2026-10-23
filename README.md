# flor2 - Local Standalone Workspace

Workspace pengujian dan pengembangan pipeline otomasi tata naskah soal ujian SD Swasta Methodist-11 Medan.

## Komponen Utama
- **Engine Editing Soal (`editor_soal.py` / `scripts/editor_soal_v1_85.py`):**
  Engine tata naskah naskah soal ujian F4 berbasis python-docx. Menerapkan standar kop master resmi, penataan seksi 2 kolom seimbang, hanging indent kelipatan 0.65 cm, opsi manual berkarakter literal tab (`\t`), dan tipografi 4-lapis Tahoma 11 pt solid black (`#000000`).
- **Auditor Naskah (`audit_naskah.py` / `scripts/audit_naskah.py`):**
  Auditor statis read-only untuk memverifikasi dokumen terhadap standar format baku SD Methodist-11 tanpa ketergantungan Word COM.
- **Template Master (`templates/SD SWASTA METHODIST.docx`):**
  Dokumen template kop master resmi ukuran Folio F4 (21.59 × 33.02 cm) dengan border ganda dan logo sekolah.
- **Automated Test Suite & Runner (`tests/`, `run_validation.py`):**
  Suite pengujian otomatis dan batch runner yang menguji naskah Kelas 1 s.d. Kelas 6 SD, memverifikasi integritas data guru (read-only SHA-256), dan membuktikan pencapaian 0 FAIL, 0 WARN.

## Cara Menjalankan

### 1. Menjalankan Unit Test & Integration Test
```bash
# Unit test engine
python -m unittest tests/test_engine.py -v

# Audit test Kelas 1–6 SD
python -m unittest tests/test_audit.py -v

# Menjalankan seluruh test suite
python -m unittest discover tests -v
```

### 2. Menjalankan Batch Validation Runner
```bash
python run_validation.py
```
Runner akan memproses sampel naskah Kelas 1 s.d. 6 SD, menjalankan audit, memverifikasi integritas data read-only, dan menghasilkan file `LAPORAN_UJI_VALIDASI.md`.

### 3. Menggunakan Engine Secara Mandiri
```bash
python editor_soal.py "path/ke/naskah_mentah_guru.docx" -o "path/ke/output_OK_edit.docx"
python audit_naskah.py "path/ke/output_OK_edit.docx"
```

## Laporan Hasil Uji
Lihat berkas [LAPORAN_UJI_VALIDASI.md](LAPORAN_UJI_VALIDASI.md) untuk rincian hasil pengujian per jenjang kelas.

