<!-- Auto-generated guidance for AI coding agents. Keep concise and actionable. -->
# Copilot instructions — projeto analise de dados

Brief: This is a small Flask web app that ingests uploaded spreadsheets, normalizes operator names, counts unique CPF/CNPJ agreements per operator/shift, and returns a generated Excel plus JSON with aggregated results.

- **Entry point:** `app.py` — single Flask app exposing:
  - `GET /` → renders `templates/index.html` (frontend with Chart.js)
  - `POST /analisar` → accepts form with `planilha` (file) and `equipe` (string). Currently only `variavel` is supported.
  - `GET /download/<file_id>` → serves the generated Excel from `temp_files/`.

- **Data flow:** upload → `read_any_file()` (multi-engine Pandas reader) → normalize names with `normalize_name()` (uses `unidecode`, uppercasing) → match against in-memory `variavel_operadores` dict → count unique `CPF/CNPJ` per operator → build Excel with `openpyxl` → save to `temp_files/` and return `file_id`.

- **Key files & dirs:**
  - `app.py` — core logic and endpoints
  - `templates/index.html` — frontend form + Chart.js rendering
  - `static/` — CSS (optional tweaks here)
  - `temp_files/` — runtime temp storage; files older than ~1 hour are removed by `limpar_temp_files()`
  - `requirements.txt` — runtime libs (Flask, pandas, openpyxl, xlrd, unidecode, lxml, html5lib)

- **Important conventions & patterns (project-specific):**
  - Names are normalized before matching: remove accents, strip, and uppercase. All operator lists in `variavel_operadores` are stored as capitalized, unaccented values — update both list and any matching logic consistently.
  - Expected input column names (Portuguese): `Nome do operador` and `CPF/CNPJ`. The app returns an error if either is missing.
  - The `equipe` form field uses the value `variavel` to select `variavel_operadores`. To add another team, add a branch in `analisar()` and a dictionary similar to `variavel_operadores`.
  - `read_any_file()` tries multiple readers in order: openpyxl (.xlsx), xlrd (.xls), pd.read_html, pd.read_csv, pd.read_table — rely on this when accepting messy inputs (HTML tables, CSVs, tab-delimited).
  - Temp files expire implicitly (cleanup called at start of `analisar()`). Downloads use `file_id` UUIDs.

- **Developer workflows / quick commands:**
  - Run locally (development):
    ```bash
    python app.py
    # app runs on http://0.0.0.0:5000 with debug=True per file
    ```
  - Test upload via curl (example):
    ```bash
    curl -F "planilha=@/path/to/file.xlsx" -F "equipe=variavel" http://localhost:5000/analisar
    ```

- **What to check when modifying code:**
  - If adding new file readers, ensure `read_any_file()` ordering doesn't break expected formats.
  - When changing operator lists, use `normalize_name()` on entries and confirm front-end labels still map to returned names (the UI expects normalized names in responses).
  - Keep `TEMP_FOLDER` consistent; do not assume persistent storage — files may be deleted after ~1 hour.
  - `xlrd==1.2.0` is required to read legacy `.xls` files (see `requirements.txt`).

- **Example edits that make sense for this repo:**
  - Add a new `equipe` (e.g., `receptivo`) — add a dict, update the branch in `analisar()`, and add option in `templates/index.html` select.
  - Improve robustness: return structured errors (JSON) consistently and add server-side logging for file read exceptions.

If any section is unclear or you want more examples (e.g., unit-test stubs, CI commands), tell me which area to expand.
