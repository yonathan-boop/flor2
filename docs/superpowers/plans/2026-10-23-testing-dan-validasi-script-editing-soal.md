# Testing dan Validasi Script Editing Soal Ujian SD Methodist-11 Implementation Plan

> **For agentic workers:** REQUIRED SUB-SKILL: Use superpowers:subagent-driven-development (recommended) or superpowers:executing-plans to implement this plan task-by-task. Steps use checkbox (`- [ ]`) syntax for tracking.

**Goal:** Setup script engine `editor_soal_v1_85.py` and auditor `audit_naskah.py` in repo `flor2`, enhance engine to achieve 100% compliance (0 FAIL, 0 WARN) on SD Kelas 1–6 with 2-column balancing and 4-layer typography, create automated test runner, and produce comprehensive test report.

**Architecture:** Port engine and auditor scripts into `flor2`, bundle official master kop template (`SD SWASTA METHODIST.docx`) for self-contained execution, fix typography leak in cloned paragraphs (Cambria -> Tahoma 11 pt) and implement 6-section continuous column balancing, build local pytest/unittest runner with read-only source assertions, and generate validation report.

**Tech Stack:** Python 3.12, `python-docx` 1.2.0, OpenXML (WML), `unittest` / automated test runner.

**Spec:** `C:\Users\yonat\OPENCODE\modules\pekerjaan_sekolah\cara_editing_soal_ujian_hermes.md`

## Global Constraints

- Standar Kertas: Folio / F4 (21.59 × 33.02 cm / 12240 × 18720 dxa)
- Margin Dokumen: Atas 0.50 cm (283 dxa), Bawah 0.60 cm (340 dxa), Kiri 1.00 cm (567 dxa), Kanan 1.00 cm (567 dxa)
- Tipografi: Tahoma 11.0 pt, Solid Black `#000000` (tanpa tema Word `themeColor`/`themeTint`)
- Indentasi Grid: Kelipatan 0.65 cm (Left 0.65 cm Hanging 0.65 cm untuk soal, Left 1.30 cm Hanging 0.65 cm untuk opsi)
- Opsi Manual: Wajib karakter literal `\t` setelah huruf opsi (`a.\t`) dan custom tab stop (1.30 cm, 4.55 cm / 5.20 cm)
- Arsitektur Seksi: 6 seksi pemutus kontinu (continuous section break) untuk keseimbangan kolom
- Integritas Data: File sumber guru dan folder `@Back UP` wajib READ-ONLY (larangan menimpa/memodifikasi)
- Target Audit: 0 FAIL, 0 WARN (100% lolos audit `audit_naskah.py`)

## Review Focus

1. Source file modification prevention: verify input files' timestamps and hashes are unchanged after processing.
2. Cloned paragraph typography leakage: math equations and drawing paragraphs must have all text runs converted to Tahoma 11 pt #000000.
3. Continuous section balancing: continuous section breaks must be present around section titles and column sections without creating blank extra pages.
4. Word COM independent fallback: engine must function reliably with pure python-docx/XML without depending on active WINWORD process.
5. Grade 1-6 edge cases: handle math fractions, low-grade lines, multi-part headers, inline options, single options.

---

### Task 1: Setup Script Files and Template into Repo

**Files:**
- Create: `templates/SD SWASTA METHODIST.docx`
- Create: `scripts/audit_naskah.py`
- Create: `scripts/editor_soal_v1_85.py`
- Create: `editor_soal.py` (root wrapper)
- Create: `audit_naskah.py` (root wrapper)

**Interfaces:**
- Produces: CLI `python editor_soal.py <input> -o <output>`, CLI `python audit_naskah.py <docx_path>`

- [ ] **Step 1: Copy official template `SD SWASTA METHODIST.docx` to `templates/`**
- [ ] **Step 2: Copy `audit_naskah.py` from `C:\Users\yonat\OPENCODE\COMMON\scripts\` to `scripts/` and root wrapper**
- [ ] **Step 3: Copy `editor_soal_v1_85.py` from `C:\Users\yonat\OPENCODE\COMMON\scripts\` to `scripts/` and root wrapper**
- [ ] **Step 4: Verify template and audit script run cleanly**
- [ ] **Step 5: Commit baseline setup**

```bash
git add templates/ scripts/ editor_soal.py audit_naskah.py
git commit -m "feat: setup editor_soal, audit_naskah, and master kop template"
```

---

### Task 2: Implement 4-Layer Typography and Section Column Architecture in Engine

**Files:**
- Modify: `scripts/editor_soal_v1_85.py`

**Interfaces:**
- Consumes: `templates/SD SWASTA METHODIST.docx`, raw docx
- Produces: Normalized docx with 0 FAIL, 0 WARN audit guarantee

- [ ] **Step 1: Write test checking typography layer on cloned drawing/equation paragraphs**
- [ ] **Step 2: Implement 4-layer typography (Lapis 1-4) in `scripts/editor_soal_v1_85.py`**
- [ ] **Step 3: Implement 6 continuous sections architecture (`w:cols w:num="2" w:space="567"` / `w:space="720"`)**
- [ ] **Step 4: Verify P0 1pt gap fix and section break integration**
- [ ] **Step 5: Run tests on sample docx and verify 0 FAIL, 0 WARN**
- [ ] **Step 6: Commit engine enhancements**

```bash
git add scripts/editor_soal_v1_85.py
git commit -m "feat(engine): implement 4-layer typography and continuous 2-column balancing"
```

---

### Task 3: Build Automated Test Suite & Local Validation Runner

**Files:**
- Create: `tests/test_engine.py`
- Create: `tests/test_audit.py`
- Create: `run_validation.py`

**Interfaces:**
- Produces: CLI `python run_validation.py` executing all Kelas 1-6 checks and outputting audit table

- [ ] **Step 1: Write `tests/test_engine.py` covering: kop cloning, font Tahoma 11pt, hanging indent, tab stop options, section columns, read-only assertion**
- [ ] **Step 2: Write `tests/test_audit.py` verifying auditor passes edited files with 0 FAIL, 0 WARN**
- [ ] **Step 3: Create `run_validation.py` batch runner across Kelas 1–6 samples**
- [ ] **Step 4: Run test suite with `python -m unittest discover tests`**
- [ ] **Step 5: Commit test suite and runner**

```bash
git add tests/ run_validation.py
git commit -m "test: add automated test suite and batch validation runner"
```

---

### Task 4: Execute Full Validation across Kelas 1–6 and Generate Test Report

**Files:**
- Create: `LAPORAN_UJI_VALIDASI.md`
- Modify: `README.md`

**Interfaces:**
- Consumes: Test runner results
- Produces: Complete markdown report `LAPORAN_UJI_VALIDASI.md`

- [ ] **Step 1: Run `python run_validation.py` on full Kelas 1–6 sample matrix**
- [ ] **Step 2: Verify all output files achieve 0 FAIL, 0 WARN in `audit_naskah.py`**
- [ ] **Step 3: Verify source files in Desktop and `@Back UP` were untouched (hashes verified)**
- [ ] **Step 4: Write comprehensive test report in `LAPORAN_UJI_VALIDASI.md`**
- [ ] **Step 5: Update `README.md` with instructions on running tests and the engine**
- [ ] **Step 6: Commit test report and docs**

```bash
git add LAPORAN_UJI_VALIDASI.md README.md
git commit -m "docs: generate comprehensive test report and update documentation"
```
