<<<<<<< HEAD
# soran-excel-extractor
Batch-extract structured data from thousands of Excel (`.xlsx`) files into flat tables ready for database ingestion.
=======
# Soran Excel Extractor

Batch-extract structured data from thousands of Excel (`.xlsx`) files into flat tables ready for database ingestion.

This project produces three datasets:

- **File Metadata**: one row per file
- **File Records**: multiple rows per file (from sheets 3+)
- **File Images**: extracted embedded images with file/sheet mapping (best-effort)

Outputs are written to both **Excel** and **CSV** (UTF-8 with BOM for Persian text).

## Key Features

- Processes many `.xlsx` files in a folder (`source/`)
- Extracts metadata from repeated semi-structured layouts (key/value patterns)
- Extracts record tables with dynamic header detection
- Starts scanning from sheet index 2 (3rd sheet onward)
- Adds `شیت` (SheetName) right after `FileID` in the records output
- Adds `نام ابزار/دقت ابزار` only for dimensional (`ابعادی`) sheets
- Extracts images from `.xlsx` internals (`xl/media`) and maps to sheets when possible
- Produces an **incomplete information report** (file + sheet + column + row)
- Per-file error handling; continues processing even if one file fails
- Final summary report printed to stdout

## Project Layout

Recommended repository layout:

- `process_excel_files.py` – main extractor
- `source/` – put your input `.xlsx` files here (not committed)
- `output/` – generated outputs (not committed)
- `scripts/build_macos.sh` – builds a macOS `.app` and `.dmg` in `dist/`

## Installation (CLI)

Create a virtual environment and install requirements:

```bash
python3 -m venv .venv
source .venv/bin/activate
pip install -r requirements.txt
```

## Run (CLI)

```bash
python3 process_excel_files.py --input "./source" --output "./output"
```

## Offline SQL Export (No Server Needed)

To generate SQL Server scripts you can run later (offline):

```bash
python3 process_excel_files.py --input "./source" --output "./output" --export-sql "./output/sql"
```

This writes:

- `schema.sql` (CREATE TABLE scripts)
- `data.sql` (INSERT statements)

During execution you should see lines like:

```text
Processing 12/1866: /path/to/file.xlsx
```

## Outputs

After a run, the `output/` folder contains:

- `file_metadata.xlsx` / `file_metadata.csv`
- `file_records.xlsx` / `file_records.csv`
- `file_images.xlsx` / `file_images.csv`
- `incomplete_information.xlsx` / `incomplete_information.csv`
- `process.log`

### Record Columns

`file_records.*` columns:

- `FileID`
- `شیت`
- `مشخصات کنترلی`
- `شرح کنترل`
- `درجه اهمیت`
- `استاندارد مرجع`
- `محدوده قابل قبول`
- `روش نمونه گیری (تناوب - تعداد) توضیحات`
- `نام ابزار/دقت ابزار` (only populated for `ابعادی` sheets)

## macOS App (.app) and DMG (.dmg)

To create a distributable macOS app bundle and DMG:

```bash
./scripts/build_macos.sh
```

Artifacts are created under:

- `dist/<APP_NAME>.app`
- `dist/<APP_NAME>.dmg`

The app runs the extractor in **Terminal** so you can watch live progress reliably on macOS systems where Tkinter/Tk can crash.

### Gatekeeper (First Launch)

If macOS blocks the app:

- Right click the app → **Open**
- Or go to **System Settings → Privacy & Security** and choose **Open Anyway**

This repo uses an ad-hoc signature for local development builds. For wide distribution, use a Developer ID and notarization.

## Notes

- This project reads cell values only; images embedded in cells do not replace cell text.
- Some workbooks report extremely large used ranges; the extractor includes an empty-row cutoff to avoid “stuck” scans.

## License

MIT. See [LICENSE](LICENSE).
>>>>>>> 44b1966 (Initial release: Soran Excel Extractor)
