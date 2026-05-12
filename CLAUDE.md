# CLAUDE.md

This file provides guidance to Claude Code (claude.ai/code) when working with code in this repository.

## What this project does

A Python CLI tool for Belgian salaried workers who need to file Brazilian income tax. It handles two distinct declaration types:

- **Carnê-Leão** (monthly): converts Belgian payslips from EUR to BRL using official exchange rates per IN SRF 208/2002.
- **DIRPF annual**: converts year-end bank account balances, capital gains from asset disposals, and cryptocurrency holdings to BRL.

## Running the tool

```bash
python main.py                          # creates input template on first run, then processes it
python main.py --input data/input.xlsx  # explicit input file
python main.py --output data/out.xlsx   # explicit output file
python main.py --year 2024              # override year (default: current year)
```

Install dependencies first:
```bash
pip install -r requirements.txt
```

There is no build step, test suite, or linter configured.

## Template

`templates/input_template.xlsx` is a committed blank workbook with all four input sheets pre-formatted (`Salarios_EUR`, `Contas_Bancarias`, `Ganhos_Capital`, `Criptomoedas`). Users copy it to `data/input_YEAR.xlsx` before filling in data. To regenerate it (e.g. after adding new sheets or columns), pick any tax year as the label-source — only the bank-account and crypto sheets embed the year in their description row:

```bash
python -c "
from spreadsheet import create_template
from assets import add_asset_sheets_to_template
YEAR = 2026
create_template('templates/input_template.xlsx', YEAR)
add_asset_sheets_to_template('templates/input_template.xlsx', YEAR)
"
```

`data/` is gitignored (personal financial data); `templates/` is committed (blank reference only).

## Architecture

Five modules with clear separation of concerns:

- **`main.py`** — CLI entry point and orchestrator. Seven numbered steps: (1) create template if input missing → (2) read salary input → (3) fetch salary rates → (4) convert → (5) write salary output → (6) `_process_assets()` for optional DIRPF asset sheets → (7) print salary summary table and Receita Federal field mapping.
- **`spreadsheet.py`** — All salary-specific Excel I/O via openpyxl. Creates `Salarios_EUR` input sheet, reads and validates it, writes `Carne_Leao_YEAR` (grouped column layout: Salario / Opcoes / Vakantiegeld / 13e Maand / Cambio / Resumo) + `Resumo_Anual` output sheets. Handles European date (DD/MM/YYYY) and number formats; tolerates older templates missing the optional stock-options / vakantiegeld / 13e maand columns.
- **`converter.py`** — Pure salary calculation using `Decimal` for precision. Applies `EUR × ECB_EUR_USD × BCB_USD_BRL = BRL` per row, producing per-income-type `rendimentos_*_brl`, `deducao_prev_*_brl`, `tributavel_*_brl`, `imposto_*_brl`, `netto_*_brl` columns plus `salario_liquido_brl` and `base_calculo_brl`. RSZ is deducted from salary and 13e maand only; not from stock options or vakantiegeld.
- **`ptax.py`** — Exchange rate fetching, caching (`cache/ptax_cache.json`), retry logic, up-to-7-day lookback for weekends/holidays. Two public functions: `get_rates()` (salary rule) and `get_spot_rates()` (spot date rule for assets).
- **`assets.py`** — All DIRPF asset logic: template sheet builders (`add_*_sheet`, plus the umbrella `add_asset_sheets_to_template`), readers, converters, and output writers for bank accounts, capital gains, and crypto.

## Two exchange rate rules

These differ and must not be confused:

- **Salary (`get_rates`)**: ECB EUR/USD on payment date; BCB USD/BRL on the last business day of the first half (days 1–15) of the **prior month** — per IN SRF 208/2002.
- **Assets (`get_spot_rates`)**: ECB EUR/USD and BCB USD/BRL both on the **target date** (with 7-day fallback). Used for: bank account year-end balance (31 Dec), capital gain transaction dates, crypto acquisition date.

## Input/output structure

**Input** (`data/input_YEAR.xlsx`) — four sheets:
- `Salarios_EUR` — one row per month, required for Carnê-Leão
- `Contas_Bancarias` — year-end EUR balances → Bens e Direitos (codes 61/62)
- `Ganhos_Capital` — asset disposals (stocks, ETFs, crypto sold) → Ganhos de Capital
- `Criptomoedas` — year-end crypto holdings at acquisition cost → Bens e Direitos código 89

**Output** (`data/output_YEAR.xlsx`) — sheets appended as data is found:
- `Carne_Leao_YEAR`, `Resumo_Anual` — always written if salary data present
- `Bens_Direitos_YEAR` — written if bank accounts filled in
- `Ganhos_Capital_YEAR` — written if capital gains filled in
- `Criptomoedas_YEAR` — written if crypto filled in

The asset sheets are optional: if a sheet is absent or all rows are empty, it is silently skipped.

## Error types

- `InputError` (spreadsheet.py) — malformed template: missing columns, duplicate months, unparseable values
- `PTAXNotFoundError` — no rate available within the 7-day fallback window
- `PTAXNetworkError` — API network/timeout failures
- `PTAXAPIError` — API response parsing failures
