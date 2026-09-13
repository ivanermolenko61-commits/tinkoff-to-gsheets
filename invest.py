"""
Tinkoff Invest → Google Sheets

Синхронизация портфеля из Tinkoff Invest API в Google Таблицу через Apps Script.

Почему два слоя:
  Google Apps Script не может напрямую обращаться к Tinkoff Invest API:
  сертификат Минцифры РФ не входит в доверенные в среде Apps Script, а
  параметр validateHttpsCertificates: false не решает проблему.
  Поэтому Python-скрипт получает данные через официальный SDK (в котором
  сертификат уже встроен) и передаёт их в Apps Script POST-запросом.

Перед запуском:
  Скопируйте .env.example в .env и заполните:
    TINKOFF_TOKEN       — токен из личного кабинета T-Invest (только чтение)
    APPS_SCRIPT_URL     — URL веб-приложения Apps Script (оканчивается на /exec)
    APPS_SCRIPT_SECRET  — секрет, совпадающий с APPS_SECRET в Script Properties
"""

import os
import time
import json as json_module
from pathlib import Path

import requests
from dotenv import load_dotenv
from t_tech.invest import Client
from t_tech.invest.utils import quotation_to_decimal

# --- Загрузка секретов ---
load_dotenv(dotenv_path=Path(__file__).parent / ".env", encoding="utf-8-sig")

TOKEN = os.environ.get("TINKOFF_TOKEN", "").strip()
APPS_SCRIPT_URL = os.environ.get("APPS_SCRIPT_URL", "").strip()
SECRET = os.environ.get("APPS_SCRIPT_SECRET", "").strip()

if not TOKEN:
    raise SystemExit("Не задан TINKOFF_TOKEN в .env")
if not APPS_SCRIPT_URL:
    raise SystemExit("Не задан APPS_SCRIPT_URL в .env")
if not SECRET:
    raise SystemExit("Не задан APPS_SCRIPT_SECRET в .env")

# --- Настройки листов ---
RAW_SHEET = 'Positions'
AGG_SHEET = 'Positions_Aggregated'
SUMMARY_BY_TYPE_SHEET = 'Positions_SummaryByType'
TYPE_SHEETS = {
    'Shares':     'Positions_Shares',
    'Bonds':      'Positions_Bonds',
    'ETFs':       'Positions_ETFs',
    'Currencies': 'Positions_Currencies',
    'Futures':    'Positions_Futures',
    'Other':      'Positions_Other',
    'Money':      'Positions_Money',
}
# Чем группировать сводку: 'name' | 'ticker' | 'figi'
AGGREGATE_BY = 'name'


class AppsScriptWriter:
    """Отправляет данные в Google Apps Script, который пишет их в таблицу."""

    def __init__(self, url, secret):
        self.url = url
        self.secret = secret

    def write_tables(self, sheets_dict):
        full_payload = {'secret': self.secret, 'sheets': sheets_dict}
        payload_str = json_module.dumps(full_payload, ensure_ascii=False)

        resp = requests.post(
            self.url,
            data={'payload': payload_str},
            timeout=180,
        )
        if resp.status_code != 200:
            raise RuntimeError(f"Apps Script вернул HTTP {resp.status_code}: {resp.text[:300]}")

        result = resp.json()
        if result.get('status') != 'ok':
            raise RuntimeError(f"Apps Script ошибка: {result.get('message')}")


def _to_str(v):
    """Аккуратно приводит enum/строку к строке в нижнем регистре."""
    if v is None:
        return ''
    s = getattr(v, 'name', v)
    return str(s).lower()


def _status_to_str(v):
    """Приводит статус (int или enum) к строке для сравнения."""
    if v is None:
        return ''
    name = getattr(v, 'name', None)
    if name:
        return str(name).upper()
    return str(v).upper()


def get_instrument_info(client, figi):
    """Возвращает справочник инструмента по FIGI или None при ошибке."""
    try:
        response = client.instruments.get_instrument_by(id_type=1, id=figi)
        instrument = response.instrument
        return {
            'figi': figi,
            'lot': instrument.lot or 1,
            'ticker': instrument.ticker or '',
            'name': instrument.name or '',
            'instrument_type': _to_str(getattr(instrument, 'instrument_type', '')),
            'currency': str(getattr(instrument, 'currency', '') or '').upper(),
        }
    except Exception as e:
        print(f"  ! Не удалось получить справочник по FIGI {figi}: {e}")
        return None


def to_category(instr_type):
    """Классификация типа инструмента → имя листа."""
    t = (instr_type or '').lower()
    if 'share' in t: return 'Shares'
    if t == 'bond': return 'Bonds'
    if t == 'etf': return 'ETFs'
    if t == 'currency': return 'Currencies'
    if t in ('future', 'futures'): return 'Futures'
    if t == 'money': return 'Money'
    return 'Other'


def aggregate_by_key(rows, key_name):
    """Агрегирует строки по ключу (name/ticker/figi) с взвешенным усреднением цен."""
    if not rows or len(rows) <= 1:
        return rows
    header = rows[0]
    H = {k: i for i, k in enumerate(header)}

    aggregated = {}
    for r in rows[1:]:
        key = str(r[H.get(key_name, 0)]).strip()
        if not key:
            continue

        if key not in aggregated:
            aggregated[key] = {
                'row': r[:],
                'types': set(), 'lots': set(), 'tickers': set(), 'currencies': set(),
                'qty_lots_sum': 0, 'qty_pcs_sum': 0, 'pos_rub_sum': 0,
                'cur_weighted_sum': 0, 'cur_qty_sum': 0,
                'avg_weighted_sum': 0, 'avg_qty_sum': 0,
            }

        ag = aggregated[key]
        ag['types'].add(r[H.get('type', 0)])
        ag['lots'].add(r[H.get('lot', 0)])
        ag['tickers'].add(r[H.get('ticker', 0)])
        ag['currencies'].add(r[H.get('instrument_currency', 0)])

        qty_lots = float(r[H.get('quantity_lots', 0)] or 0)
        qty_pcs  = float(r[H.get('quantity_pcs', 0)] or 0)
        cur_price = float(r[H.get('current_price_rub_per_piece', 0)] or 0)
        avg_price = float(r[H.get('avg_price_rub_per_piece', 0)] or 0)
        pos_rub = float(r[H.get('position_value_rub', 0)] or 0)

        ag['qty_lots_sum'] += qty_lots
        ag['qty_pcs_sum']  += qty_pcs
        ag['pos_rub_sum']  += pos_rub

        if cur_price > 0 and qty_pcs > 0:
            ag['cur_weighted_sum'] += cur_price * qty_pcs
            ag['cur_qty_sum'] += qty_pcs
        if avg_price > 0 and qty_pcs > 0:
            ag['avg_weighted_sum'] += avg_price * qty_pcs
            ag['avg_qty_sum'] += qty_pcs

    out = [header[:]]
    for key, ag in aggregated.items():
        row = ag['row'][:]
        row[H.get('accountId', 0)] = 'ALL'
        row[H.get('type', 0)] = list(ag['types'])[0] if len(ag['types']) == 1 else 'security'
        row[H.get('figi', 0)] = ''
        row[H.get('ticker', 0)] = list(ag['tickers'])[0] if len(ag['tickers']) == 1 else ''
        row[H.get('name', 0)] = key
        row[H.get('lot', 0)] = list(ag['lots'])[0] if len(ag['lots']) == 1 else ''
        row[H.get('quantity_lots', 0)] = ag['qty_lots_sum']
        row[H.get('quantity_pcs', 0)]  = ag['qty_pcs_sum']
        row[H.get('current_price_rub_per_piece', 0)] = (
            ag['cur_weighted_sum'] / ag['cur_qty_sum'] if ag['cur_qty_sum'] else 0
        )
        row[H.get('avg_price_rub_per_piece', 0)] = (
            ag['avg_weighted_sum'] / ag['avg_qty_sum'] if ag['avg_qty_sum'] else ''
        )
        row[H.get('position_value_rub', 0)] = ag['pos_rub_sum']
        row[H.get('instrument_currency', 0)] = (
            list(ag['currencies'])[0] if len(ag['currencies']) == 1 else ''
        )
        out.append(row)

    if 'position_value_rub' in H:
        pos_idx = H['position_value_rub']
        out[1:] = sorted(out[1:], key=lambda x: float(x[pos_idx] or 0), reverse=True)

    return out


def collect_positions(client):
    """Собирает позиции по всем открытым счетам. Возвращает список строк."""
    header = [
        'accountId', 'type', 'figi', 'ticker', 'name', 'lot',
        'quantity_lots', 'quantity_pcs', 'current_price_rub_per_piece',
        'avg_price_rub_per_piece', 'price_rub_per_lot', 'position_value_rub',
        'instrument_currency',
    ]
    raw_data = [header]

    accounts_response = client.users.get_accounts()
    all_accounts = accounts_response.accounts

    accounts = [
        acc for acc in all_accounts
        if 'OPEN' in _status_to_str(getattr(acc, 'status', None))
        and 'CLOSED' not in _status_to_str(getattr(acc, 'status', None))
    ]
    print(f"Открытых счетов: {len(accounts)}")

    for idx, account in enumerate(accounts, start=1):
        account_id = account.id
        print(f"  · Обработка счёта #{idx}...")

        portfolio = client.operations.get_portfolio(account_id=account_id)
        positions = portfolio.positions
        added = 0

        for position in positions:
            figi = position.figi
            if not figi:
                continue
            info = get_instrument_info(client, figi)
            if not info:
                continue

            lot = max(1, int(info['lot']))
            qty_pcs = float(quotation_to_decimal(position.quantity))
            current_price = float(quotation_to_decimal(position.current_price))
            avg_price = (
                float(quotation_to_decimal(position.average_position_price))
                if position.average_position_price else 0
            )
            qty_lots = qty_pcs / lot if lot else qty_pcs
            price_per_lot = current_price * lot
            position_rub = current_price * qty_pcs

            raw_data.append([
                account_id, info['instrument_type'], figi, info['ticker'], info['name'],
                lot, qty_lots, qty_pcs,
                current_price, avg_price if avg_price else '',
                price_per_lot, position_rub, info['currency'],
            ])
            added += 1
            time.sleep(0.05)

        # Деньги (RUB)
        pos = client.operations.get_positions(account_id=account_id)
        money_list = getattr(pos, 'money', []) or []
        for money in money_list:
            cur = str(getattr(money, 'currency', '') or '').upper()
            if cur == 'RUB':
                rub_amount = float(quotation_to_decimal(money))
                if rub_amount > 0:
                    raw_data.append([
                        account_id, 'money', '', 'RUB', 'Деньги (RUB)',
                        '', '', '', '', '', '', rub_amount, 'RUB',
                    ])
                break

        print(f"    позиций: {added}")

    return raw_data


def build_all_sheets(raw_data):
    """Формирует словарь {имя листа: строки} со всеми выходными листами."""
    header = raw_data[0]
    H = {k: i for i, k in enumerate(header)}
    sheets = {}

    sheets[RAW_SHEET] = raw_data

    if AGGREGATE_BY in ('name', 'ticker', 'figi'):
        sheets[AGG_SHEET] = aggregate_by_key(raw_data, AGGREGATE_BY)

    # Разбивка по типам
    per_type = {cat: [header[:]] for cat in TYPE_SHEETS.keys()}
    for r in raw_data[1:]:
        is_money = str(r[H['type']]).lower() == 'money'
        cat = 'Money' if is_money else to_category(r[H['type']])
        if cat in per_type:
            per_type[cat].append(r)

    for cat, sheet_name in TYPE_SHEETS.items():
        sheets[sheet_name] = aggregate_by_key(per_type[cat], AGGREGATE_BY)

    # Свод по типам
    summary = [['type', 'position_value_rub']]
    sums = {}
    for r in raw_data[1:]:
        cat = to_category(r[H['type']])
        val = float(r[H['position_value_rub']] or 0)
        sums[cat] = sums.get(cat, 0) + val
    for cat in ('Shares', 'Bonds', 'ETFs', 'Currencies', 'Futures', 'Other'):
        if cat in sums:
            summary.append([cat, sums[cat]])

    money_sum = sum(
        float(r[H['position_value_rub']] or 0)
        for r in per_type['Money'][1:]
    )
    if money_sum > 0:
        summary.append(['Money', money_sum])
    sheets[SUMMARY_BY_TYPE_SHEET] = summary

    return sheets


def main():
    print("Подключение к Tinkoff Invest API...")
    with Client(TOKEN) as client:
        raw_data = collect_positions(client)
    print(f"Собрано строк: {len(raw_data) - 1}")

    print("Подготовка листов...")
    sheets = build_all_sheets(raw_data)

    print("Отправка в Google Sheets через Apps Script...")
    writer = AppsScriptWriter(APPS_SCRIPT_URL, SECRET)
    writer.write_tables(sheets)

    print("Готово. Все листы обновлены.")


if __name__ == '__main__':
    main()