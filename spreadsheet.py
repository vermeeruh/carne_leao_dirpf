"""Input template creation, input reading, and output writing."""
import re
from datetime import date, datetime
from pathlib import Path

import pandas as pd
from openpyxl import Workbook
from openpyxl.styles import Alignment, Font, PatternFill
from openpyxl.utils import get_column_letter


class InputError(Exception):
    pass


EXPECTED_COLUMNS = [
    'Mes',
    'Data_Pagamento',
    'Salario_Bruto_EUR',
    'Previdencia_Social_EUR',
    'Imposto_Retido_Belgica_EUR',
    'Opcoes_Acoes_EUR',
    'Imposto_Retido_Opcoes_EUR',
    'Vakantiegeld_EUR',
    'Imposto_Retido_Vakantiegeld_EUR',
    'Bonus_13e_Maand_EUR',
    'Previdencia_Social_13e_Maand_EUR',
    'Imposto_Retido_13e_Maand_EUR',
    'Salario_Liquido_EUR',
]

# Columns that are optional (absent in old templates → filled with NaN)
OPTIONAL_COLUMNS = {
    'Opcoes_Acoes_EUR',
    'Imposto_Retido_Opcoes_EUR',
    'Vakantiegeld_EUR',
    'Imposto_Retido_Vakantiegeld_EUR',
    'Bonus_13e_Maand_EUR',
    'Previdencia_Social_13e_Maand_EUR',
    'Imposto_Retido_13e_Maand_EUR',
}

_HEADER_FONT  = Font(bold=True, color='FFFFFF')
_HEADER_FILL  = PatternFill(fill_type='solid', fgColor='1F4E79')
_DESC_FONT    = Font(italic=True, color='595959')
_DESC_FILL    = PatternFill(fill_type='solid', fgColor='D9E1F2')
_ALT_FILL     = PatternFill(fill_type='solid', fgColor='F2F2F2')
_CENTER       = Alignment(horizontal='center')
_BRL_FMT      = 'R$ #,##0.00'
_EUR_FMT      = '#,##0.00'
_RATE_FMT     = '0.0000'
_DATE_FMT     = 'DD/MM/YYYY'


# ---------------------------------------------------------------------------
# Template creation
# ---------------------------------------------------------------------------

def create_template(path: str, year: int):
    wb = Workbook()
    ws = wb.active
    ws.title = 'Salarios_EUR'

    headers = [
        'Mes',
        'Data_Pagamento',
        'Salario_Bruto_EUR',
        'Previdencia_Social_EUR',
        'Imposto_Retido_Belgica_EUR',
        'Opcoes_Acoes_EUR',
        'Imposto_Retido_Opcoes_EUR',
        'Vakantiegeld_EUR',
        'Imposto_Retido_Vakantiegeld_EUR',
        'Bonus_13e_Maand_EUR',
        'Previdencia_Social_13e_Maand_EUR',
        'Imposto_Retido_13e_Maand_EUR',
        'Salario_Liquido_EUR',
    ]
    descriptions = [
        'Mes (1-12)',
        'Data do pagamento (DD/MM/AAAA)',
        'Salario bruto (bruto loon) — EUR',
        'Previdencia social salario (RSZ/ONSS) — EUR',
        'Imposto retido salario (bedrijfsvoorheffing) — EUR',
        'Stock options bruto — EUR (em branco se nao houver)',
        'Imposto retido stock options — EUR (em branco se nao houver)',
        'Vakantiegeld bruto — EUR (em branco se nao houver)',
        'Imposto retido vakantiegeld — EUR (em branco se nao houver)',
        '13e maand bruto — EUR (em branco se nao houver)',
        'Previdencia social 13e maand (RSZ/ONSS) — EUR (em branco se nao houver)',
        'Imposto retido 13e maand — EUR (em branco se nao houver)',
        'Salario liquido combinado (netto loon, todos os rendimentos) — EUR (verificacao)',
    ]
    col_widths = [8, 24, 28, 32, 38, 26, 34, 26, 34, 26, 36, 34, 42]

    # Header row
    for i, (h, w) in enumerate(zip(headers, col_widths), 1):
        cell = ws.cell(row=1, column=i, value=h)
        cell.font = _HEADER_FONT
        cell.fill = _HEADER_FILL
        cell.alignment = _CENTER
        ws.column_dimensions[get_column_letter(i)].width = w

    # Description row
    for i, desc in enumerate(descriptions, 1):
        cell = ws.cell(row=2, column=i, value=desc)
        cell.font = _DESC_FONT
        cell.fill = _DESC_FILL
        cell.alignment = Alignment(horizontal='center', wrap_text=True)
    ws.row_dimensions[2].height = 30

    # Data rows: pre-fill month numbers
    n_cols = len(headers)
    for month in range(1, 13):
        row = month + 2
        ws.cell(row=row, column=1, value=month).alignment = _CENTER
        ws.cell(row=row, column=2).number_format = _DATE_FMT
        for col in range(3, n_cols + 1):
            ws.cell(row=row, column=col).number_format = _EUR_FMT
        if month % 2 == 0:
            for col in range(1, n_cols + 1):
                ws.cell(row=row, column=col).fill = _ALT_FILL

    ws.freeze_panes = 'B3'
    Path(path).parent.mkdir(parents=True, exist_ok=True)
    wb.save(path)


# ---------------------------------------------------------------------------
# Read input
# ---------------------------------------------------------------------------

def _normalize_col(name: str) -> str:
    return re.sub(r'\s+', '_', str(name).strip())


def _parse_date(val) -> date | None:
    if val is None or (isinstance(val, float) and pd.isna(val)):
        return None
    if isinstance(val, date) and not isinstance(val, datetime):
        return val
    if isinstance(val, datetime):
        return val.date()
    s = str(val).strip()
    if not s or s == 'nan':
        return None
    for fmt in ('%d/%m/%Y', '%Y-%m-%d', '%d-%m-%Y', '%m/%d/%Y'):
        try:
            return datetime.strptime(s, fmt).date()
        except ValueError:
            continue
    raise InputError(f"Cannot parse date '{val}' — use DD/MM/YYYY format")


def _parse_eur(val) -> float:
    if val is None:
        return float('nan')
    if isinstance(val, (int, float)):
        return float(val)
    s = str(val).strip()
    if not s or s == 'nan':
        return float('nan')
    # Handle European comma-decimal notation: "1.234,56" → 1234.56
    if ',' in s and '.' in s:
        s = s.replace('.', '').replace(',', '.')
    elif ',' in s:
        s = s.replace(',', '.')
    try:
        return float(s)
    except ValueError:
        raise InputError(f"Cannot parse EUR value '{val}'")


def read_input(path: str) -> pd.DataFrame:
    try:
        df = pd.read_excel(path, sheet_name='Salarios_EUR', header=0, skiprows=[1])
    except Exception as e:
        raise InputError(f'Cannot read {path}: {e}')

    df.columns = [_normalize_col(c) for c in df.columns]

    # Required columns
    required = [c for c in EXPECTED_COLUMNS if c not in OPTIONAL_COLUMNS]
    missing = [c for c in required if c not in df.columns]
    if missing:
        raise InputError(
            f'Missing columns: {missing}\nExpected: {EXPECTED_COLUMNS}'
        )

    # Optional columns — add as NaN if not present (old template)
    for col in OPTIONAL_COLUMNS:
        if col not in df.columns:
            df[col] = float('nan')

    df = df[EXPECTED_COLUMNS].copy()

    # Drop non-data rows (Total rows, notes, etc.) — keep only valid month integers 1-12
    df['Mes'] = pd.to_numeric(df['Mes'], errors='coerce')
    df = df[df['Mes'].between(1, 12)].copy()

    # Check duplicate months
    dupes = df['Mes'][df['Mes'].duplicated()].astype(int).tolist()
    if dupes:
        raise InputError(f'Duplicate entries for month(s): {dupes}')

    # Parse dates
    df['Data_Pagamento'] = df['Data_Pagamento'].apply(_parse_date)

    # Parse EUR amounts
    for col in ['Salario_Bruto_EUR', 'Previdencia_Social_EUR',
                'Imposto_Retido_Belgica_EUR', 'Opcoes_Acoes_EUR',
                'Imposto_Retido_Opcoes_EUR', 'Vakantiegeld_EUR',
                'Imposto_Retido_Vakantiegeld_EUR', 'Bonus_13e_Maand_EUR',
                'Previdencia_Social_13e_Maand_EUR', 'Imposto_Retido_13e_Maand_EUR',
                'Salario_Liquido_EUR']:
        df[col] = df[col].apply(_parse_eur)

    return df


# ---------------------------------------------------------------------------
# Write output
# ---------------------------------------------------------------------------

def write_output(df: pd.DataFrame, path: str, year: int):
    wb = Workbook()
    ws = wb.active
    ws.title = f'Carne_Leao_{year}'

    out_headers = [
        'Mes',
        'Rendimentos_Salario_BRL',
        'Rendimentos_Opcoes_BRL',
        'Rendimentos_Vakantiegeld_BRL',
        'Rendimentos_13e_Maand_BRL',
        'Deducao_Prev_Social_Salario_BRL',
        'Deducao_Prev_Social_13e_Maand_BRL',
        'Imposto_Retido_Salario_BRL',
        'Imposto_Retido_Opcoes_BRL',
        'Imposto_Retido_Vakantiegeld_BRL',
        'Imposto_Retido_13e_Maand_BRL',
        'ECB_EUR_USD',
        'BCB_USD_BRL_Compra',
        'ECB_Data',
        'BCB_Data',
        'Salario_Liquido_BRL',
        'Base_de_Calculo_BRL',
        'Observacoes',
    ]
    col_widths = [6, 28, 26, 30, 28, 34, 36, 30, 28, 32, 30, 14, 22, 14, 14, 24, 26, 42]
    col_fmts   = [
        None,
        _BRL_FMT, _BRL_FMT, _BRL_FMT, _BRL_FMT,
        _BRL_FMT, _BRL_FMT,
        _BRL_FMT, _BRL_FMT, _BRL_FMT, _BRL_FMT,
        _RATE_FMT, _RATE_FMT,
        _DATE_FMT, _DATE_FMT,
        _BRL_FMT, _BRL_FMT, None,
    ]

    for i, (h, w) in enumerate(zip(out_headers, col_widths), 1):
        cell = ws.cell(row=1, column=i, value=h)
        cell.font = _HEADER_FONT
        cell.fill = _HEADER_FILL
        cell.alignment = _CENTER
        ws.column_dimensions[get_column_letter(i)].width = w

    data_cols = [
        'Mes',
        'rendimentos_brl', 'rendimentos_opcoes_brl',
        'rendimentos_vakantiegeld_brl', 'rendimentos_13e_maand_brl',
        'deducao_prev_brl', 'deducao_prev_13e_maand_brl',
        'imposto_retido_brl', 'imposto_opcoes_brl',
        'imposto_vakantiegeld_brl', 'imposto_13e_maand_brl',
        'ecb_eur_usd', 'bcb_usd_brl', 'ecb_date', 'bcb_date',
        'salario_liquido_brl', 'base_calculo_brl', 'notes',
    ]

    for row_idx, (_, row) in enumerate(df.iterrows(), 2):
        fill = _ALT_FILL if row_idx % 2 == 0 else None
        for col_idx, (dcol, fmt) in enumerate(zip(data_cols, col_fmts), 1):
            val = row.get(dcol)
            # Convert date objects to string for date columns
            if fmt == _DATE_FMT and hasattr(val, 'strftime'):
                val = val  # openpyxl handles date objects natively
            cell = ws.cell(row=row_idx, column=col_idx, value=val)
            if fmt:
                cell.number_format = fmt
            if fill:
                cell.fill = fill

    ws.freeze_panes = 'B2'

    # Summary sheet
    ws2 = wb.create_sheet('Resumo_Anual')
    ws2['A1'] = f'Resumo Anual {year}'
    ws2['A1'].font = Font(bold=True, size=14)
    ws2.column_dimensions['A'].width = 42
    ws2.column_dimensions['B'].width = 22
    ws2.column_dimensions['C'].width = 22

    _EUR_FMT_SUM = '[$EUR] #,##0.00'

    # Column headers
    ws2.cell(row=2, column=2, value='EUR').font = Font(bold=True)
    ws2.cell(row=2, column=3, value='BRL').font = Font(bold=True)

    summary_rows = [
        (
            'Total Rendimentos Salario',
            df['Salario_Bruto_EUR'].sum(),
            df['rendimentos_brl'].sum(),
        ),
        (
            'Total Rendimentos Stock Options',
            df['Opcoes_Acoes_EUR'].sum(),
            df['rendimentos_opcoes_brl'].sum(),
        ),
        (
            'Total Rendimentos Vakantiegeld',
            df['Vakantiegeld_EUR'].sum(),
            df['rendimentos_vakantiegeld_brl'].sum(),
        ),
        (
            'Total Rendimentos 13e Maand',
            df['Bonus_13e_Maand_EUR'].sum(),
            df['rendimentos_13e_maand_brl'].sum(),
        ),
        (
            'Total Deducoes Prev. Social Salario',
            df['Previdencia_Social_EUR'].sum(),
            df['deducao_prev_brl'].sum(),
        ),
        (
            'Total Deducoes Prev. Social 13e Maand',
            df['Previdencia_Social_13e_Maand_EUR'].sum(),
            df['deducao_prev_13e_maand_brl'].sum(),
        ),
        (
            'Total Imposto Retido Salario',
            df['Imposto_Retido_Belgica_EUR'].sum(),
            df['imposto_retido_brl'].sum(),
        ),
        (
            'Total Imposto Retido Stock Options',
            df['Imposto_Retido_Opcoes_EUR'].sum(),
            df['imposto_opcoes_brl'].sum(),
        ),
        (
            'Total Imposto Retido Vakantiegeld',
            df['Imposto_Retido_Vakantiegeld_EUR'].sum(),
            df['imposto_vakantiegeld_brl'].sum(),
        ),
        (
            'Total Imposto Retido 13e Maand',
            df['Imposto_Retido_13e_Maand_EUR'].sum(),
            df['imposto_13e_maand_brl'].sum(),
        ),
        (
            'Total Base de Calculo',
            (df['Salario_Bruto_EUR'] + df['Opcoes_Acoes_EUR']
             + df['Vakantiegeld_EUR'] + df['Bonus_13e_Maand_EUR']
             - df['Previdencia_Social_EUR'] - df['Previdencia_Social_13e_Maand_EUR']).sum(),
            df['base_calculo_brl'].sum(),
        ),
    ]
    for i, (label, eur_val, brl_val) in enumerate(summary_rows, 3):
        ws2.cell(row=i, column=1, value=label).font = Font(bold=True)
        eur_cell = ws2.cell(row=i, column=2, value=eur_val)
        eur_cell.number_format = _EUR_FMT_SUM
        brl_cell = ws2.cell(row=i, column=3, value=brl_val)
        brl_cell.number_format = _BRL_FMT

    Path(path).parent.mkdir(parents=True, exist_ok=True)
    wb.save(path)
