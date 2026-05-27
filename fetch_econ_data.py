"""Fetch economic data from U.S. government APIs for CECL reports.

Sources:
  - Unemployment rate (state): U.S. Bureau of Labor Statistics, LAUS
  - Population (state): U.S. Census Bureau, ACS 1-Year Estimates
  - Bankruptcies (state): Administrative Office of the U.S. Courts
  - Foreclosures: Not available from free federal APIs
"""

import json
import logging
from datetime import datetime
from io import BytesIO

import pandas as pd
import requests

log = logging.getLogger(__name__)

_TIMEOUT = 30  # seconds

# ── State name → (FIPS code, postal abbreviation) ──────────────────
STATE_INFO = {
    'Alabama': ('01', 'AL'), 'Alaska': ('02', 'AK'), 'Arizona': ('04', 'AZ'),
    'Arkansas': ('05', 'AR'), 'California': ('06', 'CA'), 'Colorado': ('08', 'CO'),
    'Connecticut': ('09', 'CT'), 'Delaware': ('10', 'DE'),
    'District of Columbia': ('11', 'DC'),
    'Florida': ('12', 'FL'), 'Georgia': ('13', 'GA'), 'Hawaii': ('15', 'HI'),
    'Idaho': ('16', 'ID'), 'Illinois': ('17', 'IL'), 'Indiana': ('18', 'IN'),
    'Iowa': ('19', 'IA'), 'Kansas': ('20', 'KS'), 'Kentucky': ('21', 'KY'),
    'Louisiana': ('22', 'LA'), 'Maine': ('23', 'ME'), 'Maryland': ('24', 'MD'),
    'Massachusetts': ('25', 'MA'), 'Michigan': ('26', 'MI'),
    'Minnesota': ('27', 'MN'), 'Mississippi': ('28', 'MS'),
    'Missouri': ('29', 'MO'), 'Montana': ('30', 'MT'), 'Nebraska': ('31', 'NE'),
    'Nevada': ('32', 'NV'), 'New Hampshire': ('33', 'NH'),
    'New Jersey': ('34', 'NJ'), 'New Mexico': ('35', 'NM'),
    'New York': ('36', 'NY'), 'North Carolina': ('37', 'NC'),
    'North Dakota': ('38', 'ND'), 'Ohio': ('39', 'OH'), 'Oklahoma': ('40', 'OK'),
    'Oregon': ('41', 'OR'), 'Pennsylvania': ('42', 'PA'),
    'Rhode Island': ('44', 'RI'), 'South Carolina': ('45', 'SC'),
    'South Dakota': ('46', 'SD'), 'Tennessee': ('47', 'TN'),
    'Texas': ('48', 'TX'), 'Utah': ('49', 'UT'), 'Vermont': ('50', 'VT'),
    'Virginia': ('51', 'VA'), 'Washington': ('53', 'WA'),
    'West Virginia': ('54', 'WV'), 'Wisconsin': ('55', 'WI'),
    'Wyoming': ('56', 'WY'),
}


def _lookup_state(state_name):
    """Return (fips, abbreviation) for a state name, case-insensitive."""
    for name, info in STATE_INFO.items():
        if name.lower() == state_name.strip().lower():
            return info
    return None, None


# ── BLS Local Area Unemployment Statistics ─────────────────────────

def _fetch_unemployment(state_fips):
    """Fetch state unemployment rate from BLS LAUS API (v1, no key required).

    Returns (rate_as_decimal, year) or (None, None).
    """
    series_id = f'LASST{state_fips}0000000000003'
    url = 'https://api.bls.gov/publicAPI/v1/timeseries/data/'
    payload = json.dumps({'seriesid': [series_id]})
    headers = {'Content-type': 'application/json'}

    resp = requests.post(url, data=payload, headers=headers, timeout=_TIMEOUT)
    resp.raise_for_status()
    data = resp.json()

    if data.get('status') != 'REQUEST_SUCCEEDED':
        raise ValueError(f"BLS API error: {data.get('message', 'unknown')}")

    series = data['Results']['series'][0]['data']

    # Prefer annual average (period M13)
    for item in series:
        if item['period'] == 'M13':
            return float(item['value']) / 100.0, int(item['year'])

    # Fallback: most recent December value
    for item in series:
        if item['period'] == 'M12':
            return float(item['value']) / 100.0, int(item['year'])

    # Fallback: most recent month available
    if series:
        return float(series[0]['value']) / 100.0, int(series[0]['year'])

    return None, None


# ── Census Bureau Population ───────────────────────────────────────

def _fetch_population(state_fips):
    """Fetch state population from Census ACS 1-Year Estimates.

    Returns (population_int, data_year) or (None, None).
    """
    now_year = datetime.now().year
    # ACS 1-year data is released ~Sep of the following year,
    # so check from (current_year - 1) backwards.
    for year in range(now_year - 1, now_year - 4, -1):
        try:
            url = (f'https://api.census.gov/data/{year}/acs/acs1'
                   f'?get=B01003_001E,NAME&for=state:{state_fips}')
            resp = requests.get(url, timeout=_TIMEOUT)
            if resp.status_code == 200:
                data = resp.json()
                if len(data) > 1:
                    return int(data[1][0]), year
        except Exception:
            continue

    return None, None


# ── US Courts Bankruptcy Filings ───────────────────────────────────

def _fetch_bankruptcies(state_abbrev):
    """Fetch total bankruptcy filings for a state from US Courts Table F-2.

    Tries fiscal-year (ending Sep 30) and calendar-year tables.
    Multi-district states are summed automatically.
    Returns (total_filings, period_description) or (None, None).
    """
    now_year = datetime.now().year
    st = state_abbrev.upper()

    # Try fiscal-year endings first (published sooner), then calendar year
    for dt_suffix, label_fmt in [('0930', '12 months ending 9/30'),
                                  ('0630', '12 months ending 6/30'),
                                  ('0331', '12 months ending 3/31'),
                                  ('1231', 'calendar year')]:
        for year in range(now_year, now_year - 3, -1):
            try:
                url = (f'https://www.uscourts.gov/sites/default/files/'
                       f'data_tables/bf_f2_{dt_suffix}.{year}.xlsx')
                resp = requests.get(url, timeout=_TIMEOUT, allow_redirects=True)
                if resp.status_code != 200 or len(resp.content) < 1000:
                    continue

                df = pd.read_excel(BytesIO(resp.content), header=None)
                total = 0
                found = False
                for _, row in df.iterrows():
                    district = str(row.iloc[0]).strip()
                    # Match exact state or multi-district (e.g. NY,N / NY,E)
                    if district == st or district.startswith(st + ','):
                        val = row.iloc[1]  # "Total All Chapters" column
                        if pd.notna(val):
                            try:
                                total += int(val)
                                found = True
                            except (ValueError, TypeError):
                                pass

                if found:
                    period = f'{label_fmt}/{year}'
                    return total, period

            except Exception:
                continue

    return None, None


# ── Public entry point ─────────────────────────────────────────────

def fetch_economic_data(state_name, county_name):
    """Fetch economic indicators from government APIs.

    Parameters
    ----------
    state_name : str   Full state name, e.g. "Connecticut".
    county_name : str  County name, e.g. "Hartford".

    Returns
    -------
    dict with keys:
        unemployment_rate  (decimal, e.g. 0.042)
        population         (int)
        bankruptcies       (int)
        foreclosures       (not fetched – remains from WARM/config)
        _sources           dict mapping field name → source citation string
    """
    state_fips, state_abbrev = _lookup_state(state_name)
    if not state_fips:
        log.warning(f'  Unknown state "{state_name}" – skipping external fetch')
        return {}

    result = {'state': state_name, 'county': county_name}
    sources = {}

    # ── Unemployment rate ──
    try:
        rate, data_year = _fetch_unemployment(state_fips)
        if rate is not None:
            result['unemployment_rate'] = rate
            sources['unemployment_rate'] = (
                f'U.S. Bureau of Labor Statistics, Local Area Unemployment '
                f'Statistics (LAUS), {data_year} Annual Average')
            log.info(f'    Unemployment: {rate*100:.1f}% ({data_year})')
    except Exception as e:
        log.warning(f'    BLS unemployment fetch failed: {e}')

    # ── Population ──
    try:
        pop, data_year = _fetch_population(state_fips)
        if pop is not None:
            result['population'] = pop
            sources['population'] = (
                f'U.S. Census Bureau, American Community Survey '
                f'1-Year Estimates ({data_year})')
            log.info(f'    Population: {pop:,} ({data_year})')
    except Exception as e:
        log.warning(f'    Census population fetch failed: {e}')

    # ── Bankruptcies ──
    try:
        bk, period = _fetch_bankruptcies(state_abbrev)
        if bk is not None:
            result['bankruptcies'] = bk
            sources['bankruptcies'] = (
                f'Administrative Office of the U.S. Courts, '
                f'Table F-2 ({period})')
            log.info(f'    Bankruptcies: {bk:,} ({period})')
    except Exception as e:
        log.warning(f'    US Courts bankruptcy fetch failed: {e}')

    # ── Foreclosures ──
    # No free federal API provides state-level foreclosure filing counts.
    # The value from the WARM file / config is retained.
    sources['foreclosures'] = 'Institution records / WARM file'

    result['_sources'] = sources
    return result
