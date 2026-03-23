# TFM management

Automates creation of per-student defense folders for TFM/TFT evaluation:
- matches students across ZIP submissions, Excel master list, and director reports,
- generates editable actas,
- packages each student dossier ready to upload (for example, to OneDrive).

## No-Install Mode (Recommended for End Users)

You can run this tool without installing Python, pip, or creating environments.

Use a prebuilt executable:
- **Windows**: `crear carpetas TFM.exe` (also distributed as `tfm_assigner.exe`)
- **Linux/macOS (Unix)**: `crear carpetas TFM` (also distributed as `tfm_assigner`)

How to use:
1. Download the executable for your OS from the project release/artifacts.
2. Copy that executable into your run folder (the folder that contains ZIP/Excel/template/reports).
3. Run:

Windows:
```powershell
.\crear carpetas TFM.exe
```

Unix (Linux/macOS):
```bash
chmod +x "crear carpetas TFM"
./"crear carpetas TFM"
```

The executable behaves exactly like `python tfm_folders.py`.

## What This Script Produces

After running, the tool creates:

`output/`
- `estudiantes/Apellido1 Apellido2, Nombre/`
  - manuscript (copied as received from ZIP)
  - `informe_director_Apellido1_Apellido2_Nombre.pdf` (if found)
  - `diapositivas_Apellido1_Apellido2_Nombre.pdf` (if found)
  - generated acta PDF (prefilled student data, tribunal fields remain editable)
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
mamba create -n tfm_assigner -y python=3.11 pandas openpyxl rapidfuzz unidecode pypdf
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

## Build the Executable (for maintainers)

### Unix (Linux/macOS)

```bash
./scripts/build_executable.sh
```

### Windows (PowerShell)

```powershell
.\scripts\build_executable.ps1
```

Binary output:
- Unix: `dist/crear carpetas TFM` and `dist/tfm_assigner`
- Windows: `dist/crear carpetas TFM.exe` and `dist/tfm_assigner.exe`

### Automated cross-platform builds (CI)

GitHub Actions workflow:
- `.github/workflows/build-executables.yml`

It builds one-file executables for:
- Linux
- macOS
- Windows

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
  --ambiguity-margin 4 \
  --only-student "Apellido"
```

## Required Input Files

Run the script from a workspace folder (or pass it with `--workspace`).

Required:
1. ZIP of submitted manuscripts (`.zip`) in workspace root.
2. Excel master list (`.xlsx`) with student data (workspace root).
3. Acta template (`.pdf`) fillable PDF template (workspace root).
Optional:
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
   Use the official fillable PDF acta model used by your program.

## Expected File Behavior and Detection Rules

### ZIP
- Any `.zip` filename is accepted.
- If multiple ZIP files exist in workspace root, the newest by modification time is used.
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
- A fillable `.pdf` template in workspace root is required.
- Keep only one real acta template `.pdf` in the run folder to avoid ambiguity.
- Template filename controls acta output naming. Use:
  - `<codigo_asignatura>_<Apellido1 Apellido2, Nombre>_<edicion>.pdf`
  - Example template name:
    - `50789_Apellido 1 Apellido 2, Nombre_Abril 25.pdf`
  - Generated actas will follow exactly:
    - `50789_<Apellido real del estudiante, Nombre>_Abril 25.pdf`
  - `edicion` is taken from Excel metadata when available; if missing, it is inferred from template filename.

## Manuscript Naming

For each extracted manuscript:
1. Leading numeric prefix is removed from filename:
   - `2891870066 - Eliana Yurany Santos Santana - tfm_final_4.docx`
   - becomes `Eliana Yurany Santos Santana - tfm_final_4.docx`
2. Manuscripts are kept in original format (no office conversion dependency).

## Folder Structure You Must Respect

Recommended run folder layout before execution:

```text
run_folder/
  entregas_tfm_abril.zip
  2510_ListadoTFT_MBIF.xlsx
  50789_Apellido 1 Apellido 2, Nombre_Abril 25.pdf
  informes/
    Informe del director ... .pdf
    Informe del director ... .pdf
```

Do not move ZIP/Excel/template out of workspace root unless you also change the code.

## What Can Break Detection (Read This)

These changes commonly break or degrade the process:

1. Missing Excel columns for student name or DNI.
2. Director reports only available as scanned PDFs without OCR text.
3. No valid fillable PDF acta template in root.
4. Multiple competing PDF templates in root.
5. PDF template without AcroForm fillable fields.
6. Running with a wrong `--workspace` folder.

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
