# TFM management

Automates creation of per-student defense folders for TFM/TFT evaluation:
- matches students across ZIP submissions, Excel master list, and director reports,
- generates editable actas,
- packages each student dossier ready to upload (for example, to OneDrive).

## What This Script Produces

After running, the tool creates:

`output/`
- `estudiantes/Apellido1 Apellido2, Nombre/`
  - manuscript (PDF; non-PDF manuscripts are converted to PDF when possible)
  - matched director report PDF (if found)
  - generated acta DOCX
- `resumen_procesamiento.csv`
- `revision_manual.csv`
- `logs/procesamiento_YYYYMMDD_HHMMSS.log`

## Core Processing Rule

Only students with a manuscript inside the ZIP are processed.

The Excel is used as master metadata source, but does not decide who gets processed.

## Quick Start (Clone + Install + Run)

### 1. Clone repository

```bash
git clone https://github.com/MiquelSendra/tfm.git
cd tfm
```

### 2. Create a reproducible environment (recommended: mamba)

```bash
mamba create -n tfm_assigner -y python=3.11 pandas openpyxl rapidfuzz unidecode pypdf python-docx
mamba activate tfm_assigner
```

### 3. Alternative install (pip + requirements.txt)

If you already have an environment, install dependencies with:

```bash
pip install -r requirements.txt
```

### 4. Run

```bash
python tfm_folders.py
```

Or with explicit workspace:

```bash
python tfm_folders.py --workspace /path/to/run_folder
```

## Environment

Use your existing mamba environment:

```bash
mamba activate tfm_assigner
```

Run from the repository root:

```bash
python tfm_folders.py
```

Optional CLI parameters:

```bash
python tfm_folders.py \
  --workspace /path/to/run_folder \
  --output-dir-name output \
  --match-threshold 82 \
  --ambiguity-margin 4
```

## Required Input Files

Run the script from a workspace folder (or pass it with `--workspace`).

Required:
1. ZIP of submitted manuscripts with numeric filename only, like `182571167.zip` (workspace root).
2. Excel master list (`.xlsx`) with student data (workspace root).
3. Acta template (`.docx`) editable template (workspace root).
4. Director report PDFs (`.pdf`) anywhere under workspace (root or subfolders such as `informes/`).

Important:
- PDF discovery is recursive.
- Folders/files under hidden paths and `output/` are ignored to avoid reprocessing generated artifacts.

## How To Obtain Each Input

1. ZIP submissions:
   Download the ZIP export from the submission platform/LMS for the call.
2. Excel master list:
   Export or obtain the official roster spreadsheet from your academic/admin source.
3. Director reports:
   Download all director report PDFs for the same call.
4. Acta template:
   Use the official editable DOCX acta model used by your program.

## Expected File Behavior and Detection Rules

### ZIP
- Must match regex: `^\d+\.zip$`.
- If multiple numeric ZIP files exist, the newest by modification time is used.
- Manuscript files accepted inside ZIP: `.pdf`, `.doc`, `.docx`, `.odt`.

### Excel
- First `.xlsx` found by newest modification time is used.
- Header row is auto-detected (it does not need to start at row 1).
- Metadata fields `TÍTULO:` and `EDICIÓN:` are parsed from sheet content.

### Director reports (PDF)
- Filename is ignored for identity.
- Student name is extracted from PDF text content.
- Matching strategy:
  1. full-name fuzzy match (primary),
  2. surname-only fallback only when unique and unambiguous in Excel.
- Ambiguous surname fallback is sent to `revision_manual.csv` (no auto-assignment).
- Reports must be valid text PDFs (scanned image PDFs without OCR can fail).

### Acta template
- `.docx` templates are inspected; the one that best matches expected markers is selected.
- Keep only one real acta template `.docx` in the run folder to avoid ambiguity.

## Manuscript Naming and Conversion

For each extracted manuscript:
1. Leading numeric prefix is removed from filename:
   - `2891870066 - Eliana Yurany Santos Santana - tfm_final_4.docx`
   - becomes `Eliana Yurany Santos Santana - tfm_final_4.docx`
2. If manuscript is not PDF, script attempts conversion to PDF via LibreOffice headless.

LibreOffice binary detection order:
1. `LIBREOFFICE_BIN` environment variable (full path),
2. `soffice` in `PATH`,
3. `libreoffice` in `PATH`,
4. `/Applications/LibreOffice.app/Contents/MacOS/soffice` (macOS default).

If no LibreOffice binary is found, conversion fails for non-PDF manuscripts and the issue is logged in:
- `output/resumen_procesamiento.csv` (`manuscript_status`, `notes`)
- `output/logs/...`

To make DOCX/ODT manuscript conversion work, install LibreOffice and make `soffice` available in `PATH`, or set:

```bash
export LIBREOFFICE_BIN="/full/path/to/soffice"
```

## Folder Structure You Must Respect

Recommended run folder layout before execution:

```text
run_folder/
  182571167.zip
  2510_ListadoTFT_MBIF.xlsx
  50789_Apellido 1 Apellido 2, Nombre_Abril 25_NUEVA.docx
  informes/
    Informe del director ... .pdf
    Informe del director ... .pdf
```

Do not move ZIP/Excel/template out of workspace root unless you also change the code.

## What Can Break Detection (Read This)

These changes commonly break or degrade the process:

1. ZIP filename not numeric (example: `entregas.zip`).
2. Missing Excel columns for student name or DNI.
3. Director reports only available as scanned PDFs without OCR text.
4. No valid DOCX acta template in root.
5. Multiple competing templates with similar content in root.
6. Non-PDF manuscript with no LibreOffice available.
7. Running with a wrong `--workspace` folder.

## Output Files You Should Review Every Run

1. `output/resumen_procesamiento.csv`:
   final record per processed student, output paths, statuses, and notes.
2. `output/revision_manual.csv`:
   ambiguous/unmatched cases that need human validation.
3. `output/logs/procesamiento_*.log`:
   detailed technical trace.

## Notes for Operators

- Re-running overwrites/regenerates outputs in `output/`.
- If previous runs left stale files, clear `output/` before a clean batch run.
- For production runs, keep one batch per workspace folder.

## License

MIT. See `LICENSE`.
