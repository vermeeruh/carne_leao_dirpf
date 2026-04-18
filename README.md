# Carnê-Leão — Belgian salary to Brazilian tax declaration

A Python CLI tool for **Belgian-based salaried workers** who must file Brazilian income tax. It converts EUR payslip data to BRL using official exchange rates and produces a ready-to-use spreadsheet for two separate declarations:

- **Carnê-Leão** (monthly): salary income, stock options, vakantiegeld, 13e maand
- **DIRPF annual**: Belgian bank account balances, capital gains, and cryptocurrency holdings

## Requirements

- Python 3.11+
- Dependencies: `pandas`, `openpyxl`, `requests`

```bash
pip install -r requirements.txt
```

## Quickstart

A blank template is included in the repo at `templates/input_template.xlsx`. Copy it to the `data/` folder and rename it before filling in your data:

```bash
cp templates/input_template.xlsx data/input_2024.xlsx
# Fill in data/input_2024.xlsx, then run:
python main.py --year 2024
```

Alternatively, let the tool generate the template for you on first run:

```bash
python main.py --year 2024   # exits after creating data/input_2024.xlsx
# Fill in data/input_2024.xlsx, then run again:
python main.py --year 2024
```

Output is written to `data/output_YEAR.xlsx`. Exchange rates are cached in `cache/ptax_cache.json` so subsequent runs are instant.

> **Note:** `data/` is excluded from git to protect personal financial data. Only the blank template in `templates/` is committed.

## Input spreadsheet

The template has four sheets. All sheets except `Salarios_EUR` are optional — leave them empty or unfilled to skip.

### `Salarios_EUR` — monthly salary (required for Carnê-Leão)

| Column | Description |
|---|---|
| `Mes` | Month number 1–12 (pre-filled) |
| `Data_Pagamento` | Date salary was credited — DD/MM/YYYY |
| `Salario_Bruto_EUR` | Bruto loon |
| `Previdencia_Social_EUR` | RSZ/ONSS contribution |
| `Imposto_Retido_Belgica_EUR` | Bedrijfsvoorheffing |
| `Opcoes_Acoes_EUR` | Stock options gross value at exercise (optional) |
| `Imposto_Retido_Opcoes_EUR` | Tax withheld on stock options (optional) |
| `Vakantiegeld_EUR` | Vakantiegeld gross (optional) |
| `Imposto_Retido_Vakantiegeld_EUR` | Tax withheld on vakantiegeld (optional) |
| `Bonus_13e_Maand_EUR` | 13e maand gross (optional) |
| `Previdencia_Social_13e_Maand_EUR` | RSZ/ONSS on 13e maand (optional) |
| `Imposto_Retido_13e_Maand_EUR` | Tax withheld on 13e maand (optional) |
| `Salario_Liquido_EUR` | Netto loon — for self-verification only |

Leave rows blank for months with no salary.

### `Contas_Bancarias` — Belgian bank accounts (DIRPF Bens e Direitos)

Year-end EUR balance for each account, converted to BRL using the 31 Dec exchange rate.
In DIRPF: declare under **Bens e Direitos**, código 61 (conta corrente) or 62 (poupança).

| Column | Description |
|---|---|
| `Banco` | Bank name (e.g. ING Belgium) |
| `IBAN` | Account IBAN |
| `Descricao` | Description for your DIRPF |
| `Saldo_EUR` | Balance at 31 Dec — EUR |

### `Ganhos_Capital` — capital gains (DIRPF Ganhos de Capital)

One row per disposed asset (stocks, ETFs, crypto sold during the year). The BRL cost and proceeds are calculated at the rates on each respective transaction date.

| Column | Description |
|---|---|
| `Descricao` | Asset description (e.g. Ações Apple Inc.) |
| `Data_Aquisicao` | Acquisition date — DD/MM/YYYY |
| `Custo_EUR` | Total acquisition cost — EUR |
| `Data_Alienacao` | Disposal/sale date — DD/MM/YYYY |
| `Receita_EUR` | Total sale proceeds — EUR |

For **crypto disposals**, add them here too (with the ticker in the description).

### `Criptomoedas` — crypto holdings at year-end (DIRPF Bens e Direitos)

Year-end crypto balances declared at **acquisition cost** in BRL (Brazilian rule — not market value).
In DIRPF: declare under **Bens e Direitos**, código 89 (criptoativos).

| Column | Description |
|---|---|
| `Nome` | Cryptocurrency name (e.g. Bitcoin) |
| `Ticker` | Symbol (e.g. BTC) |
| `Quantidade` | Quantity held at 31 Dec |
| `Custo_Aquisicao_EUR` | Total acquisition cost — EUR |
| `Data_Aquisicao` | Approximate acquisition date — DD/MM/YYYY |

## Output spreadsheet

| Sheet | Content |
|---|---|
| `Carne_Leao_YEAR` | Monthly BRL values with exchange rates — enter into Carnê-Leão website |
| `Resumo_Anual` | Annual totals in EUR and BRL |
| `Bens_Direitos_YEAR` | Bank account BRL values for DIRPF |
| `Ganhos_Capital_YEAR` | Per-asset BRL gain/loss for DIRPF |
| `Criptomoedas_YEAR` | Crypto BRL cost basis for DIRPF |

## Exchange rate methodology

All conversions follow the two-step formula: **EUR × ECB EUR/USD × BCB USD/BRL compra = BRL**

| Declaration | ECB EUR/USD | BCB USD/BRL |
|---|---|---|
| Salary (Carnê-Leão) | Payment date | Last business day of first half of prior month — per IN SRF 208/2002 |
| Bank accounts | 31 Dec | 31 Dec |
| Capital gains | Transaction date | Transaction date |
| Crypto holdings | Acquisition date | Acquisition date |

If a rate is unavailable on the target date (weekend/holiday), the tool falls back to the nearest prior business day, up to 7 days back. All fetched rates are cached locally.

## Command-line options

```
python main.py [--year YEAR] [--input FILE] [--output FILE]
```

| Option | Default | Description |
|---|---|---|
| `--year YEAR` | Current calendar year | Tax year to process. Controls default file names and year-specific labels (e.g. balance date in bank account sheet). |
| `--input FILE` | `data/input_YEAR.xlsx` | Path to the filled-in input spreadsheet. |
| `--output FILE` | `data/output_YEAR.xlsx` | Path where the output spreadsheet is written. Created or overwritten on each run. |

**Examples**

```bash
# Process the current year using default file names
python main.py

# Process a specific past year
python main.py --year 2023

# Use custom file paths (useful when managing multiple scenarios)
python main.py --year 2024 --input ~/tax/belgium_2024.xlsx --output ~/tax/output_2024.xlsx
```

**First-run behaviour**

If `--input` does not exist, the tool creates a blank template at that path and exits. Fill it in and re-run. Alternatively, copy `templates/input_template.xlsx` to `data/input_YEAR.xlsx` and fill it in before the first run.

## Data sources

- **ECB EUR/USD**: [ECB Data Portal](https://data-api.ecb.europa.eu) — daily reference rates
- **BCB USD/BRL**: [BCB PTAX API](https://olinda.bcb.gov.br) — Fechamento PTAX (compra)
