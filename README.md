# Inventory Control Centre

A Streamlit and SQLite application built from `Inventory list(1).xlsx`. It keeps the original records as individual inventory lots, lets users maintain stock and expiry information, and produces downloadable reports.

## Included features

- Dashboard with inventory, expiry, reorder, value and data-quality indicators
- Searchable inventory list with filters and a bulk quick editor
- Add, edit, archive and restore inventory records
- Stock receipts, issues and adjustments with an auditable movement history
- Separate stock and expiry statuses so missing quantities are not incorrectly treated as zero
- Full inventory, reorder, expiry, valuation, data-quality and movement reports
- Excel and CSV report downloads
- Excel/CSV import in append or confirmed replacement mode
- Downloadable import template and SQLite database backup
- Configurable business name, currency symbol and expiry-warning window

## What was learned from the workbook

The source contains one visible sheet, `Inventory List`, with 407 populated inventory rows. It mainly describes laboratory reagents and consumables. Common brands include AGAPPE, LEADS, CORAL, REAGENT, CPC and BIOSYSTEM. The useful source fields are item, brand, test volume, description, unit price, total effective tests, selling price per test, stock quantity, reorder values, opened date and expiry date.

The workbook also contains historical duplicates, hidden rows, incomplete stock values and inconsistent text in some date cells. The application therefore:

- preserves every populated source row;
- assigns a unique application inventory code while retaining the source ID and row number;
- imports valid dates and stores unparsed source values as review notes;
- calculates cost per test, profit per test and inventory value only when the required inputs exist;
- flags missing stock, price, reorder and expiry information for review.

## Windows quick start

1. Install Python 3.10 or newer from the official Python website. During installation, enable **Add Python to PATH**.
2. Extract this project ZIP.
3. Double-click `run_app.bat`.
4. The first run creates a local virtual environment and installs the required packages.
5. Open `http://localhost:8501` if the browser does not open automatically.

The first application start imports all 407 records from `data/Inventory_list_source.xlsx` into `data/inventory.db`.

## Manual start

```bash
python -m venv .venv

# Windows
.venv\Scripts\activate

# macOS/Linux
source .venv/bin/activate

python -m pip install -r requirements.txt
python -m streamlit run app.py
```

## Database and backup

The app uses SQLite, so no separate database server is required. Its working database is `data/inventory.db`. Use **Import & Backup → Download database backup** before replacing the current inventory or moving to a new computer.

To store the database elsewhere, set the environment variable `INVENTORY_DB_PATH` to the required `.db` path before starting Streamlit.

## Import rules

- `Item Name` (or `Item`) is required.
- Inventory codes are generated when blank.
- Dates can use `YYYY-MM-DD` or `DD-MM-YYYY`.
- Blank numeric cells remain unknown; they are not changed to zero.
- Append mode adds records to the current list.
- Replace mode clears inventory and movement history only after confirmation.

Use the downloadable template from the app for the most reliable mapping.

## Validation

Run the included validation without starting Streamlit:

```bash
python smoke_test.py
```

It verifies the 407-row workbook import, unique inventory codes, SQLite CRUD operations, stock receipts/issues, status calculations and Excel report generation.

## Project files

- `app.py` — Streamlit interface
- `database.py` — SQLite schema, CRUD, calculations and movement transactions
- `importer.py` — original-workbook and Excel/CSV import logic
- `reports.py` — report filters and formatted Excel/CSV generation
- `data/Inventory_list_source.xlsx` — unchanged source workbook used for first-run import
- `smoke_test.py` — repeatable functional validation
