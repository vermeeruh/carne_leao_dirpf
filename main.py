"""
Carnê-Leão — Belgian salary EUR→BRL converter + DIRPF asset declarations

Usage:
  python main.py                   # first run: creates input template
  python main.py                   # subsequent runs: reads input, writes output
  python main.py --input my.xlsx   # custom input path
  python main.py --output out.xlsx # custom output path
  python main.py --year 2024       # tax year (default: current year)
"""
import argparse
import sys
from datetime import datetime
from pathlib import Path

from assets import (
    add_asset_sheets_to_template,
    convert_bank_accounts,
    convert_capital_gains,
    convert_crypto,
    read_bank_accounts,
    read_capital_gains,
    read_crypto,
    write_assets_output,
)
from converter import apply_conversions
from ptax import PTAXAPIError, PTAXNetworkError, PTAXNotFoundError, get_rates
from spreadsheet import InputError, create_template, read_input, write_output

BASE_DIR   = Path(__file__).parent
DATA_DIR   = BASE_DIR / 'data'
CACHE_DIR  = BASE_DIR / 'cache'
CACHE_FILE = CACHE_DIR / 'ptax_cache.json'


def main():
    parser = argparse.ArgumentParser(
        description='Belgian payslip EUR→BRL converter for Carnê-Leão and DIRPF'
    )
    parser.add_argument('--input',  default=None,
                        help='Path to input spreadsheet (default: data/input_YEAR.xlsx)')
    parser.add_argument('--output', default=None,
                        help='Path for output spreadsheet (default: data/output_YEAR.xlsx)')
    parser.add_argument('--year',   default=datetime.now().year, type=int,
                        help='Tax year (default: current year)')
    args = parser.parse_args()

    if args.input is None:
        args.input = str(DATA_DIR / f'input_{args.year}.xlsx')
    if args.output is None:
        args.output = str(DATA_DIR / f'output_{args.year}.xlsx')

    DATA_DIR.mkdir(exist_ok=True)
    CACHE_DIR.mkdir(exist_ok=True)

    # Step 1 — create template if input does not exist
    if not Path(args.input).exists():
        create_template(args.input, args.year)
        add_asset_sheets_to_template(args.input, args.year)
        print(f'\nTemplate created: {args.input}')
        print('\nSheet "Salarios_EUR" — fill one row per month:')
        print('  - Data_Pagamento: date salary was credited (DD/MM/YYYY)')
        print('  - Salario_Bruto_EUR: bruto loon')
        print('  - Previdencia_Social_EUR: RSZ/ONSS contribution')
        print('  - Imposto_Retido_Belgica_EUR: bedrijfsvoorheffing')
        print('  - Salario_Liquido_EUR: netto loon (for self-check)')
        print('\nSheet "Contas_Bancarias" (optional) — Belgian bank account year-end balances')
        print('Sheet "Ganhos_Capital" (optional) — stocks, ETFs, crypto disposals')
        print('Sheet "Criptomoedas" (optional) — crypto holdings at year-end (at cost)')
        print('\nRe-run after filling in data.')
        sys.exit(0)

    # Step 2 — read salary input
    try:
        df = read_input(args.input)
    except InputError as e:
        print(f'ERROR: {e}')
        sys.exit(1)

    has_data = df['Salario_Bruto_EUR'].notna()
    working  = df[has_data].copy()
    skipped  = df[~has_data]['Mes'].dropna().astype(int).tolist()

    if skipped:
        print(f'Note: No salary data for month(s) {skipped} — skipping.')

    if working.empty:
        print('No salary rows found in input. Fill in the template and re-run.')
        sys.exit(1)

    missing_dates = working[working['Data_Pagamento'].isna()]['Mes'].astype(int).tolist()
    if missing_dates:
        print(f'ERROR: Missing payment date (Data_Pagamento) for month(s): {missing_dates}')
        sys.exit(1)

    for _, row in working.iterrows():
        if row['Data_Pagamento'].year != args.year:
            print(
                f"Warning: Month {int(row['Mes'])} payment date "
                f"{row['Data_Pagamento']} is outside {args.year} — verify this is correct."
            )

    # Step 3 — fetch exchange rates for salary
    print(f'\nFetching salary exchange rates for {len(working)} month(s)...')
    ecb_rates, ecb_dates, bcb_rates, bcb_dates, notes = [], [], [], [], []

    for _, row in working.iterrows():
        month        = int(row['Mes'])
        payment_date = row['Data_Pagamento']
        print(f'  Month {month:02d} ({payment_date}) ... ', end='', flush=True)
        try:
            r = get_rates(payment_date, str(CACHE_FILE))
        except PTAXNetworkError as e:
            print(f'\nERROR: Network error for month {month}: {e}')
            sys.exit(1)
        except PTAXNotFoundError as e:
            print(f'\nERROR: Rate not found for month {month}: {e}')
            sys.exit(1)
        except PTAXAPIError as e:
            print(f'\nERROR: API error for month {month}: {e}')
            sys.exit(1)

        ecb_rates.append(r['ecb_eur_usd'])
        ecb_dates.append(r['ecb_date'])
        bcb_rates.append(r['bcb_usd_brl'])
        bcb_dates.append(r['bcb_date'])
        notes.append(r['notes'])
        print(
            f"ECB {r['ecb_eur_usd']:.4f} EUR/USD (date: {r['ecb_date']})  |  "
            f"BCB {r['bcb_usd_brl']:.4f} USD/BRL compra (date: {r['bcb_date']})"
            + (f'  [{r["notes"]}]' if r['notes'] != 'OK' else '')
        )

    working['ecb_eur_usd'] = ecb_rates
    working['ecb_date']    = ecb_dates
    working['bcb_usd_brl'] = bcb_rates
    working['bcb_date']    = bcb_dates
    working['notes']       = notes

    # Step 4 — convert salary
    working = apply_conversions(working)

    # Step 5 — write salary output
    try:
        write_output(working, args.output, args.year)
    except PermissionError:
        print(f'\nERROR: Cannot write to {args.output}')
        print('Close the file in Excel and re-run.')
        sys.exit(1)

    # Step 6 — process asset declarations (bank accounts, capital gains, crypto)
    _process_assets(args.input, args.output, args.year)

    # Step 7 — print salary summary
    sep = '-' * 110
    print(f'\nOutput written to: {args.output}')
    print()
    print(
        f"{'Mes':>4}  {'Salario':>12}  {'Opcoes':>12}  {'Vakant.':>12}  {'13e Mnd':>12}  "
        f"{'DedSal':>10}  {'Ded13':>10}  {'ECB':>8}  {'BCB':>8}"
    )
    print(sep)
    for _, row in working.iterrows():
        flag = f'  [{row["notes"]}]' if row['notes'] != 'OK' else ''
        print(
            f"{int(row['Mes']):>4}  "
            f"{row['rendimentos_brl']:>12,.2f}  "
            f"{row['rendimentos_opcoes_brl']:>12,.2f}  "
            f"{row['rendimentos_vakantiegeld_brl']:>12,.2f}  "
            f"{row['rendimentos_13e_maand_brl']:>12,.2f}  "
            f"{row['deducao_prev_brl']:>10,.2f}  "
            f"{row['deducao_prev_13e_maand_brl']:>10,.2f}  "
            f"{row['ecb_eur_usd']:>8.4f}  {row['bcb_usd_brl']:>8.4f}{flag}"
        )
    print(sep)
    print(
        f"{'TOTAL':>4}  "
        f"{working['rendimentos_brl'].sum():>12,.2f}  "
        f"{working['rendimentos_opcoes_brl'].sum():>12,.2f}  "
        f"{working['rendimentos_vakantiegeld_brl'].sum():>12,.2f}  "
        f"{working['rendimentos_13e_maand_brl'].sum():>12,.2f}  "
        f"{working['deducao_prev_brl'].sum():>10,.2f}  "
        f"{working['deducao_prev_13e_maand_brl'].sum():>10,.2f}"
    )
    print()
    print(
        'Carnê-Leão — enter BRL values into the Receita Federal website:\n'
        '  Rendimentos_Salario_BRL              = Rendimentos recebidos / Trabalho assalariado (salary)\n'
        '  Rendimentos_Opcoes_BRL               = Rendimentos recebidos / Trabalho assalariado (stock options)\n'
        '  Rendimentos_Vakantiegeld_BRL         = Rendimentos recebidos / Trabalho assalariado (vakantiegeld)\n'
        '  Rendimentos_13e_Maand_BRL            = Rendimentos recebidos / Trabalho assalariado (13e maand)\n'
        '  Deducao_Prev_Social_Salario_BRL      = Deducoes / Previdencia Social estrangeira (salary)\n'
        '  Deducao_Prev_Social_13e_Maand_BRL    = Deducoes / Previdencia Social estrangeira (13e maand)\n'
        '  Imposto_Retido_Salario_BRL           = Imposto retido na fonte no exterior (salary)\n'
        '  Imposto_Retido_Opcoes_BRL            = Imposto retido na fonte no exterior (stock options)\n'
        '  Imposto_Retido_Vakantiegeld_BRL      = Imposto retido na fonte no exterior (vakantiegeld)\n'
        '  Imposto_Retido_13e_Maand_BRL         = Imposto retido na fonte no exterior (13e maand)'
    )


def _process_assets(input_path: str, output_path: str, year: int):
    """Read, convert, and append asset declaration sheets if data is present."""
    df_banks  = read_bank_accounts(input_path)
    df_gains  = read_capital_gains(input_path)
    df_crypto = read_crypto(input_path)

    if df_banks is None and df_gains is None and df_crypto is None:
        return

    print('\n--- Asset declarations ---')

    if df_banks is not None:
        print(f'\nFetching 31 Dec {year} rates for {len(df_banks)} bank account(s)...')
        try:
            df_banks = convert_bank_accounts(df_banks, year, str(CACHE_FILE))
        except (PTAXNetworkError, PTAXNotFoundError, PTAXAPIError) as e:
            print(f'ERROR fetching bank account rates: {e}')
            df_banks = None
        else:
            print(
                f"  ECB {df_banks['ECB_EUR_USD'].iloc[0]:.4f} EUR/USD  |  "
                f"BCB {df_banks['BCB_USD_BRL'].iloc[0]:.4f} USD/BRL"
                + (f"  [{df_banks['Observacoes'].iloc[0]}]"
                   if df_banks['Observacoes'].iloc[0] != 'OK' else '')
            )
            for _, row in df_banks.iterrows():
                banco = row.get('Banco') or '—'
                iban  = row.get('IBAN') or '—'
                print(f"  {banco} ({iban}): EUR {row['Saldo_EUR']:,.2f} → R$ {row['Saldo_BRL']:,.2f}")
            print(
                '\nDIRPF — Bens e Direitos (foreign bank accounts):\n'
                '  Código 61 = conta corrente / Código 62 = conta poupança\n'
                '  Situação em 31/12: use Saldo_BRL from sheet Bens_Direitos_YEAR'
            )

    if df_gains is not None:
        print(f'\nFetching rates for {len(df_gains)} capital gain transaction(s)...')
        try:
            df_gains = convert_capital_gains(df_gains, str(CACHE_FILE))
        except (PTAXNetworkError, PTAXNotFoundError, PTAXAPIError) as e:
            print(f'ERROR fetching capital gain rates: {e}')
            df_gains = None
        else:
            for _, row in df_gains.iterrows():
                flag = f'  [{row["Observacoes"]}]' if row['Observacoes'] != 'OK' else ''
                print(
                    f"  {row['Descricao']}: "
                    f"Custo R$ {row['Custo_BRL']:,.2f} → Receita R$ {row['Receita_BRL']:,.2f} "
                    f"→ Ganho R$ {row['Ganho_BRL']:,.2f}{flag}"
                )
            total_ganho = df_gains['Ganho_BRL'].sum()
            print(f"  TOTAL GANHO: R$ {total_ganho:,.2f}")
            print(
                '\nDIRPF — Ganhos de Capital no Exterior:\n'
                '  See sheet Ganhos_Capital_YEAR for per-asset breakdown.\n'
                '  Report each disposal in Renda Variável / Ganhos de Capital.'
            )

    if df_crypto is not None:
        print(f'\nFetching acquisition date rates for {len(df_crypto)} crypto holding(s)...')
        try:
            df_crypto = convert_crypto(df_crypto, str(CACHE_FILE))
        except (PTAXNetworkError, PTAXNotFoundError, PTAXAPIError) as e:
            print(f'ERROR fetching crypto rates: {e}')
            df_crypto = None
        else:
            for _, row in df_crypto.iterrows():
                ticker = row.get('Ticker') or ''
                label  = f"{row['Nome']} ({ticker})" if ticker else row['Nome']
                flag   = f'  [{row["Observacoes"]}]' if row['Observacoes'] != 'OK' else ''
                print(f"  {label}: EUR {row['Custo_Aquisicao_EUR']:,.2f} → R$ {row['Custo_BRL']:,.2f}{flag}")
            print(
                '\nDIRPF — Bens e Direitos (criptoativos):\n'
                '  Código 89 — use Custo_BRL (acquisition cost, not market value).\n'
                '  Crypto disposals during the year go in Ganhos_Capital sheet instead.'
            )

    if any(x is not None for x in (df_banks, df_gains, df_crypto)):
        try:
            write_assets_output(output_path, year, df_banks, df_gains, df_crypto)
        except PermissionError:
            print(f'\nERROR: Cannot write asset sheets to {output_path}')
            print('Close the file in Excel and re-run.')


if __name__ == '__main__':
    main()
