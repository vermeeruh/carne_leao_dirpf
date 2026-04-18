"""
Exchange rate fetching for Carnê-Leão foreign income conversion.

Per IN SRF 208/2002 and Receita Federal guidance:
  - Salaried income: EUR -> USD (ECB rate on payment date)
                     USD -> BRL (BCB 'dólar compra' PTAX from last business
                                day of first half of prior month)
"""
import json
import time
from datetime import date, timedelta
from pathlib import Path

import requests


class PTAXNotFoundError(Exception):
    pass


class PTAXNetworkError(Exception):
    pass


class PTAXAPIError(Exception):
    pass


# ---------------------------------------------------------------------------
# Cache helpers
# ---------------------------------------------------------------------------

def _load_cache(cache_path: str) -> dict:
    p = Path(cache_path)
    if not p.exists():
        return {}
    try:
        return json.loads(p.read_text(encoding='utf-8'))
    except json.JSONDecodeError:
        print(f'Warning: cache file {cache_path} is corrupted — starting fresh.')
        return {}


def _save_cache(cache_path: str, cache: dict):
    Path(cache_path).write_text(
        json.dumps(cache, indent=2, default=str), encoding='utf-8'
    )


# ---------------------------------------------------------------------------
# BCB: USD/BRL dólar compra
# ---------------------------------------------------------------------------

def _fetch_bcb_usd_brl(
    target_date: date, cache: dict, cache_path: str, max_lookback: int = 7
) -> tuple:
    """Return (rate, effective_date, note) for BCB dólar compra PTAX."""
    for delta in range(max_lookback + 1):
        d = target_date - timedelta(days=delta)
        key = f'BCB_USD_BRL_{d.isoformat()}'
        if key in cache:
            note = 'OK' if delta == 0 else f'BCB fallback: {target_date} -> {d}'
            return cache[key], d, note

        date_str = d.strftime('%m-%d-%Y')
        url = (
            'https://olinda.bcb.gov.br/olinda/servico/PTAX/versao/v1/odata/'
            f"CotacaoMoedaDia(moeda=@moeda,dataCotacao=@dataCotacao)"
            f"?@moeda='USD'&@dataCotacao='{date_str}'&$format=json"
        )
        try:
            resp = _get_with_retry(url)
        except PTAXNetworkError:
            if delta < max_lookback:
                continue
            raise

        data = resp.json().get('value', [])
        fechamento = [x for x in data if x.get('tipoBoletim') == 'Fechamento PTAX']
        if not fechamento:
            intermediario = [x for x in data if 'Intermediário' in x.get('tipoBoletim', '')]
            if intermediario:
                fechamento = [sorted(intermediario, key=lambda x: x['dataHoraCotacao'])[-1]]

        if fechamento:
            rate = float(fechamento[0]['cotacaoCompra'])
            cache[key] = rate
            _save_cache(cache_path, cache)
            note = 'OK' if delta == 0 else f'BCB fallback: {target_date} -> {d}'
            return rate, d, note

    raise PTAXNotFoundError(
        f'No BCB USD/BRL rate found within {max_lookback} days before {target_date}'
    )


# ---------------------------------------------------------------------------
# ECB: EUR/USD
# ---------------------------------------------------------------------------

def _fetch_ecb_eur_usd(
    target_date: date, cache: dict, cache_path: str, max_lookback: int = 7
) -> tuple:
    """Return (rate, effective_date, note) for ECB EUR/USD reference rate."""
    for delta in range(max_lookback + 1):
        d = target_date - timedelta(days=delta)
        key = f'ECB_EUR_USD_{d.isoformat()}'
        if key in cache:
            note = 'OK' if delta == 0 else f'ECB fallback: {target_date} -> {d}'
            return cache[key], d, note

        url = (
            'https://data-api.ecb.europa.eu/service/data/EXR/D.USD.EUR.SP00.A'
            f'?startPeriod={d.isoformat()}&endPeriod={d.isoformat()}&format=jsondata'
        )
        try:
            resp = _get_with_retry(url, headers={'Accept': 'application/json'})
        except PTAXNetworkError:
            if delta < max_lookback:
                continue
            raise

        if resp.status_code == 404:
            continue  # no data for this date (weekend/holiday)

        try:
            resp.raise_for_status()
            jdata = resp.json()
            obs = jdata['dataSets'][0]['series']['0:0:0:0:0']['observations']
            last_key = sorted(obs.keys(), key=int)[-1]
            rate = float(obs[last_key][0])
            cache[key] = rate
            _save_cache(cache_path, cache)
            note = 'OK' if delta == 0 else f'ECB fallback: {target_date} -> {d}'
            return rate, d, note
        except (KeyError, IndexError, ValueError):
            continue

    raise PTAXNotFoundError(
        f'No ECB EUR/USD rate found within {max_lookback} days before {target_date}'
    )


# ---------------------------------------------------------------------------
# Date helpers
# ---------------------------------------------------------------------------

def _last_business_day_first_half(payment_date: date) -> date:
    """
    Return the last business day (Mon–Fri) on or before the 15th of the month
    prior to payment_date. This is the BCB rate date per IN SRF 208/2002.
    """
    if payment_date.month == 1:
        prior_15 = date(payment_date.year - 1, 12, 15)
    else:
        prior_15 = date(payment_date.year, payment_date.month - 1, 15)

    d = prior_15
    while d.weekday() >= 5:  # Saturday=5, Sunday=6
        d -= timedelta(days=1)
    return d


# ---------------------------------------------------------------------------
# HTTP helper
# ---------------------------------------------------------------------------

def _get_with_retry(url: str, headers: dict = None, timeout: int = 10):
    kwargs = {'timeout': timeout}
    if headers:
        kwargs['headers'] = headers
    try:
        return requests.get(url, **kwargs)
    except requests.Timeout:
        time.sleep(3)
        try:
            return requests.get(url, **kwargs)
        except requests.Timeout as e:
            raise PTAXNetworkError(f'Timeout fetching {url}: {e}')
    except requests.RequestException as e:
        raise PTAXNetworkError(f'Network error: {e}')


# ---------------------------------------------------------------------------
# Public API
# ---------------------------------------------------------------------------

def get_spot_rates(target_date: date, cache_path: str) -> dict:
    """
    Fetch ECB EUR/USD and BCB USD/BRL spot rates for a specific date.

    Used for asset valuations (bank accounts, capital gains, crypto) where
    the BCB rate on the actual date applies — unlike get_rates(), which uses
    the salary-specific prior-month first-half rule from IN SRF 208/2002.

    Returns the same dict structure as get_rates().
    """
    cache = _load_cache(cache_path)
    ecb_rate, ecb_date, ecb_note = _fetch_ecb_eur_usd(target_date, cache, cache_path)
    bcb_rate, bcb_date, bcb_note = _fetch_bcb_usd_brl(target_date, cache, cache_path)
    parts = [n for n in (ecb_note, bcb_note) if n != 'OK']
    return {
        'ecb_eur_usd': ecb_rate,
        'ecb_date':    ecb_date,
        'bcb_usd_brl': bcb_rate,
        'bcb_date':    bcb_date,
        'notes':       '; '.join(parts) if parts else 'OK',
    }


def get_rates(payment_date: date, cache_path: str) -> dict:
    """
    Fetch both ECB EUR/USD and BCB USD/BRL rates for the given payment date.

    Returns:
        {
          'ecb_eur_usd': float,
          'ecb_date': date,
          'bcb_usd_brl': float,
          'bcb_date': date,
          'notes': str,   # 'OK' or description of any fallbacks used
        }
    """
    cache = _load_cache(cache_path)

    ecb_rate, ecb_date, ecb_note = _fetch_ecb_eur_usd(payment_date, cache, cache_path)

    bcb_target = _last_business_day_first_half(payment_date)
    bcb_rate, bcb_date, bcb_note = _fetch_bcb_usd_brl(bcb_target, cache, cache_path)

    parts = [n for n in (ecb_note, bcb_note) if n != 'OK']
    return {
        'ecb_eur_usd': ecb_rate,
        'ecb_date':    ecb_date,
        'bcb_usd_brl': bcb_rate,
        'bcb_date':    bcb_date,
        'notes':       '; '.join(parts) if parts else 'OK',
    }
