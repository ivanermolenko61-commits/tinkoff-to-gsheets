"""
Tinkoff Invest → Google Sheets

Собирает позиции, операции и снимок портфеля по всем открытым счетам
и отправляет их в Apps Script, который раскладывает данные по листам.

Создаваемые листы:
  Positions, Positions_Aggregated, Positions_Shares, Positions_Bonds,
  Positions_ETFs, Positions_Currencies, Positions_Futures, Positions_Other,
  Positions_Money, Positions_SummaryByType, Operations, YearReport,
  PortfolioSnapshots (накапливается, не перезаписывается).

Секреты берутся из:
  1) Переменных окружения (GitHub Actions, CI)
  2) Файла .env рядом со скриптом (локальный запуск)
  Приоритет у переменных окружения — они НЕ перезаписываются файлом .env.
"""

import os
import time
import json as json_module
from collections import Counter
from datetime import datetime, timedelta, timezone
from pathlib import Path

import requests
from dotenv import load_dotenv
from t_tech.invest import Client
from t_tech.invest.utils import quotation_to_decimal

# Загружаем .env, только если файл есть и переменные ещё не установлены
# (load_dotenv по умолчанию НЕ перезаписывает существующие переменные окружения).
env_path = Path(__file__).parent / ".env"
if env_path.exists():
    load_dotenv(dotenv_path=env_path, encoding="utf-8-sig")

TOKEN = os.environ.get("TINKOFF_TOKEN", "").strip()
APPS_SCRIPT_URL = os.environ.get("APPS_SCRIPT_URL", "").strip()
SECRET = os.environ.get("APPS_SCRIPT_SECRET", "").strip()

if not TOKEN:
    raise SystemExit("Не задан TINKOFF_TOKEN (проверьте .env или переменные окружения)")
if not APPS_SCRIPT_URL:
    raise SystemExit("Не задан APPS_SCRIPT_URL (проверьте .env или переменные окружения)")
if not SECRET:
    raise SystemExit("Не задан APPS_SCRIPT_SECRET (проверьте .env или переменные окружения)")

OPERATIONS_DAYS = 1095
REPORT_WINDOW_DAYS = 365
SNAPSHOT_DAYS_AGO = 365

RAW_SHEET = 'Positions'
AGG_SHEET = 'Positions_Aggregated'
SUMMARY_BY_TYPE_SHEET = 'Positions_SummaryByType'
OPERATIONS_SHEET = 'Operations'
YEAR_REPORT_SHEET = 'YearReport'
SNAPSHOTS_SHEET = 'PortfolioSnapshots'

TYPE_SHEETS = {
    'Shares':     'Positions_Shares',
    'Bonds':      'Positions_Bonds',
    'ETFs':       'Positions_ETFs',
    'Currencies': 'Positions_Currencies',
    'Futures':    'Positions_Futures',
    'Other':      'Positions_Other',
    'Money':      'Positions_Money',
}
AGGREGATE_BY = 'name'

DUPLICATE_MONEY_TICKERS = {'RUB000UTSTOM'}


# ----------------------------------------------------------------------------
# Классификация операций
# ----------------------------------------------------------------------------
def categorize_operation(op_type):
    t = (op_type or '').upper()
    if 'OPERATION_TYPE' in t:
        if 'TAX' in t: return 'tax'
        if 'FEE' in t: return 'fee'
        if 'DIVIDEND' in t: return 'dividend'
        if 'COUPON' in t: return 'coupon'
        if 'INPUT' in t: return 'input'
        if 'OUTPUT' in t: return 'output'
        if 'BUY' in t: return 'buy'
        if 'SELL' in t or 'REPAYMENT' in t: return 'sell'
        return 'other'
    if 'НАЛОГ' in t: return 'tax'
    if 'КОМИССИ' in t: return 'fee'
    if 'ДИВИДЕНД' in t: return 'dividend'
    if 'КУПОН' in t: return 'coupon'
    if 'ПОПОЛНЕНИЕ' in t or 'ВВОД' in t: return 'input'
    if 'ВЫВОД' in t: return 'output'
    if 'ПОКУПКА' in t: return 'buy'
    if 'ПРОДАЖА' in t: return 'sell'
    if 'ПОГАШЕНИЕ' in t: return 'sell'
    return 'other'


# ----------------------------------------------------------------------------
# HTTP-клиент Apps Script
# ----------------------------------------------------------------------------
class AppsScriptClient:
    def __init__(self, url, secret):
        self.url = url
        self.secret = secret

    def write_tables(self, sheets_dict):
        full_payload = {'secret': self.secret, 'sheets': sheets_dict}
        payload_str = json_module.dumps(full_payload, ensure_ascii=False)
        resp = requests.post(self.url, data={'payload': payload_str}, timeout=300)
        print("HTTP статус:", resp.status_code)
        print("Ответ Apps Script:", resp.text[:800])
        if resp.status_code != 200:
            raise RuntimeError(f"Apps Script HTTP {resp.status_code}: {resp.text[:300]}")
        result = resp.json()
        if result.get('status') != 'ok':
            raise RuntimeError(f"Apps Script ошибка: {result.get('message')}")
        return result

    def get_snapshot_n_days_ago(self, days_ago):
        try:
            resp = requests.get(
                self.url,
                params={
                    'secret': self.secret,
                    'action': 'get_snapshot_n_days_ago',
                    'days_ago': days_ago,
                },
                timeout=30,
            )
            if resp.status_code != 200:
                print(f"  ! HTTP {resp.status_code}")
                return None
            data = resp.json()
            if data.get('status') != 'ok':
                print(f"  ! Ошибка: {data.get('message')}")
                return None
            return data.get('snapshot')
        except Exception as e:
            print(f"  ! Ошибка запроса снимка: {e}")
            return None


# ----------------------------------------------------------------------------
# Вспомогательные функции
# ----------------------------------------------------------------------------
def _to_str(v):
    if v is None: return ''
    return str(getattr(v, 'name', v)).lower()


def _status_to_str(v):
    if v is None: return ''
    name = getattr(v, 'name', None)
    return str(name).upper() if name else str(v).upper()


def get_instrument_info(client, figi):
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
        print(f"  ! Справочник по {figi}: {e}")
        return None


def to_category(instr_type):
    t = (instr_type or '').lower()
    if 'share' in t: return 'Shares'
    if t == 'bond': return 'Bonds'
    if t == 'etf': return 'ETFs'
    if t == 'currency': return 'Currencies'
    if t in ('future', 'futures'): return 'Futures'
    if t == 'money': return 'Money'
    return 'Other'


def _parse_datetime(s):
    if not s: return None
    try:
        return datetime.strptime(str(s), '%Y-%m-%d %H:%M:%S')
    except ValueError:
        pass
    try:
        return datetime.strptime(str(s)[:10], '%Y-%m-%d')
    except ValueError:
        return None


# ----------------------------------------------------------------------------
# XIRR
# ----------------------------------------------------------------------------
def _xirr(cashflows):
    if not cashflows or len(cashflows) < 2:
        return None
    cashflows = sorted(cashflows, key=lambda x: x[0])
    d0 = cashflows[0][0]

    def npv(rate):
        total = 0.0
        for d, amt in cashflows:
            days = (d - d0).days
            total += amt / ((1.0 + rate) ** (days / 365.0))
        return total

    low, high = -0.999, 10.0
    npv_low, npv_high = npv(low), npv(high)
    if npv_low * npv_high > 0:
        return None
    for _ in range(300):
        mid = (low + high) / 2.0
        npv_mid = npv(mid)
        if abs(npv_mid) < 1e-6:
            break
        if npv_low * npv_mid < 0:
            high = mid
        else:
            low = mid
            npv_low = npv_mid
    return ((low + high) / 2.0) * 100.0


# ----------------------------------------------------------------------------
# Агрегация
# ----------------------------------------------------------------------------
def aggregate_by_key(rows, key_name):
    if not rows or len(rows) <= 1: return rows
    header = rows[0]
    H = {k: i for i, k in enumerate(header)}
    aggregated = {}
    for r in rows[1:]:
        key = str(r[H.get(key_name, 0)]).strip()
        if not key: continue
        if key not in aggregated:
            aggregated[key] = {
                'row': r[:], 'types': set(), 'lots': set(),
                'tickers': set(), 'currencies': set(),
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


# ----------------------------------------------------------------------------
# Сбор позиций
# ----------------------------------------------------------------------------
def collect_positions(client, accounts):
    header = [
        'accountId', 'type', 'figi', 'ticker', 'name', 'lot',
        'quantity_lots', 'quantity_pcs', 'current_price_rub_per_piece',
        'avg_price_rub_per_piece', 'price_rub_per_lot', 'position_value_rub',
        'instrument_currency',
    ]
    raw_data = [header]
    skipped_money_dupes = 0

    for idx, account in enumerate(accounts, start=1):
        account_id = account.id
        print(f"  · Обработка счёта #{idx}...")
        portfolio = client.operations.get_portfolio(account_id=account_id)
        added = 0

        for position in portfolio.positions:
            figi = position.figi
            if not figi:
                continue
            info = get_instrument_info(client, figi)
            if not info:
                continue

            ticker = info['ticker']
            if ticker in DUPLICATE_MONEY_TICKERS:
                skipped_money_dupes += 1
                continue

            lot = max(1, int(info['lot']))
            qty_pcs = float(quotation_to_decimal(position.quantity))
            current_price = float(quotation_to_decimal(position.current_price))
            avg_price = (
                float(quotation_to_decimal(position.average_position_price))
                if position.average_position_price else 0
            )

            nkd_per_unit = 0.0
            nkd_field = getattr(position, 'current_nkd', None)
            if nkd_field:
                nkd_per_unit = float(quotation_to_decimal(nkd_field))

            qty_lots = qty_pcs / lot if lot else qty_pcs
            price_per_lot = current_price * lot
            position_rub = (current_price + nkd_per_unit) * qty_pcs
            current_price_with_nkd = current_price + nkd_per_unit

            raw_data.append([
                account_id, info['instrument_type'], figi, ticker, info['name'],
                lot, qty_lots, qty_pcs,
                current_price_with_nkd, avg_price if avg_price else '',
                price_per_lot, position_rub, info['currency'],
            ])
            added += 1
            time.sleep(0.05)

        pos = client.operations.get_positions(account_id=account_id)
        for money in (getattr(pos, 'money', []) or []):
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

    if skipped_money_dupes:
        print(f"  ℹ️ Пропущено денежных дубликатов (RUB000UTSTOM): {skipped_money_dupes}")

    return raw_data


# ----------------------------------------------------------------------------
# Сбор операций
# ----------------------------------------------------------------------------
def fetch_operations_chunk(client, account_id, from_dt, to_dt):
    try:
        resp = client.operations.get_operations(account_id=account_id, from_=from_dt, to=to_dt)
        return list(resp.operations or [])
    except Exception as e:
        print(f"      ! Ошибка {from_dt.date()}..{to_dt.date()}: {e}")
        return []


def fetch_operations_for_account(client, account_id, from_date, to_date):
    all_ops = []
    current = from_date
    while current < to_date:
        chunk_end = min(current + timedelta(days=90), to_date)
        all_ops.extend(fetch_operations_chunk(client, account_id, current, chunk_end))
        current = chunk_end
    return all_ops


def build_operations_data(client, accounts, from_date, to_date):
    header = [
        'date', 'accountId', 'type', 'ticker', 'name',
        'quantity', 'price_rub', 'payment_rub', 'currency',
        'figi', 'operation_id',
    ]
    rows = [header]
    instr_cache = {}
    type_counter = Counter()
    for idx, account in enumerate(accounts, start=1):
        account_id = account.id
        print(f"  · Операции по счёту #{idx} ({account_id})...")
        ops = fetch_operations_for_account(client, account_id, from_date, to_date)
        print(f"    найдено операций: {len(ops)}")
        for op in ops:
            figi = getattr(op, 'figi', '') or ''
            ticker = name = ''
            if figi:
                if figi not in instr_cache:
                    instr_cache[figi] = get_instrument_info(client, figi) or {}
                info = instr_cache[figi]
                ticker = info.get('ticker', '')
                name = info.get('name', '')
            dt = getattr(op, 'date', None)
            date_str = dt.strftime('%Y-%m-%d %H:%M:%S') if dt else ''
            op_type = getattr(op, 'type', '') or getattr(op, 'operation_type', '')
            op_type_str = str(getattr(op_type, 'name', op_type) or '')
            type_counter[op_type_str] += 1
            quantity = float(getattr(op, 'quantity', 0) or 0)
            price = float(quotation_to_decimal(getattr(op, 'price', None))) if getattr(op, 'price', None) else 0.0
            payment = float(quotation_to_decimal(getattr(op, 'payment', None))) if getattr(op, 'payment', None) else 0.0
            currency = str(getattr(op, 'currency', '') or '').upper()
            op_id = str(getattr(op, 'id', '') or '')
            rows.append([
                date_str, account_id, op_type_str, ticker, name,
                quantity, price, payment, currency, figi, op_id,
            ])
        time.sleep(0.5)
    if len(rows) > 1:
        rows[1:] = sorted(rows[1:], key=lambda r: r[0], reverse=True)
    print("\n  Уникальные типы операций:")
    for op_type, count in type_counter.most_common():
        print(f"    {op_type:55} {count:>4} шт → {categorize_operation(op_type)}")
    return rows


# ----------------------------------------------------------------------------
# Снимок портфеля
# ----------------------------------------------------------------------------
def build_today_snapshot(raw_positions, today=None):
    if today is None:
        today = datetime.now(timezone.utc).replace(tzinfo=None)
    H = {k: i for i, k in enumerate(raw_positions[0])}
    portfolio_value = 0.0
    money_value = 0.0
    for r in raw_positions[1:]:
        ptype = str(r[H['type']] or '').lower()
        v = float(r[H['position_value_rub']] or 0)
        if ptype == 'money':
            money_value += v
        else:
            portfolio_value += v
    total = portfolio_value + money_value
    return [
        ['date', 'total_value', 'portfolio_value', 'money_value'],
        [today.strftime('%Y-%m-%d'), round(total, 2),
         round(portfolio_value, 2), round(money_value, 2)],
    ]


# ----------------------------------------------------------------------------
# Годовой отчёт
# ----------------------------------------------------------------------------
def build_year_report(operations_rows, raw_positions, snapshot_365, today=None):
    if today is None:
        today = datetime.now(timezone.utc).replace(tzinfo=None)

    H_ops = {k: i for i, k in enumerate(operations_rows[0])}
    H_pos = {k: i for i, k in enumerate(raw_positions[0])}
    period_start = today - timedelta(days=REPORT_WINDOW_DAYS)

    year_dividends = year_coupons = year_fees = year_taxes = 0.0
    year_inputs = year_outputs = 0.0

    xirr_flows_all = []
    total_inputs = total_outputs = 0.0
    total_dividends = total_coupons = total_fees = total_taxes = 0.0

    for r in operations_rows[1:]:
        op_type = str(r[H_ops['type']] or '')
        category = categorize_operation(op_type)
        try:
            payment = float(r[H_ops['payment_rub']] or 0)
        except (ValueError, TypeError):
            payment = 0.0
        dt = _parse_datetime(r[H_ops['date']])
        in_period = dt is not None and dt >= period_start

        if category == 'dividend':
            total_dividends += payment
            if in_period: year_dividends += payment
        elif category == 'coupon':
            total_coupons += payment
            if in_period: year_coupons += payment
        elif category == 'fee':
            total_fees += payment
            if in_period: year_fees += payment
        elif category == 'tax':
            total_taxes += payment
            if in_period: year_taxes += payment
        elif category == 'input':
            total_inputs += payment
            if in_period: year_inputs += payment
            if dt is not None and payment > 0:
                xirr_flows_all.append((dt, -payment))
        elif category == 'output':
            total_outputs += abs(payment)
            if in_period: year_outputs += abs(payment)
            if dt is not None and payment < 0:
                xirr_flows_all.append((dt, -payment))

    portfolio_value = money_value = 0.0
    for r in raw_positions[1:]:
        ptype = str(r[H_pos['type']] or '').lower()
        v = float(r[H_pos['position_value_rub']] or 0)
        if ptype == 'money':
            money_value += v
        else:
            portfolio_value += v
    total_value = portfolio_value + money_value

    xirr_all_time = _xirr(xirr_flows_all + [(today, total_value)]) if xirr_flows_all else None

    start_value = None
    start_date_str = None
    actual_days = None
    simple_period_pct = None
    simple_period_rub = None
    xirr_period = None

    if snapshot_365 and snapshot_365.get('date'):
        start_dt = _parse_datetime(snapshot_365['date'])
        if start_dt:
            start_value = float(snapshot_365.get('total_value') or 0)
            start_date_str = snapshot_365['date']
            actual_days = (today - start_dt).days

            simple_period_rub = total_value - start_value - year_inputs + year_outputs
            if start_value > 0:
                simple_period_pct = simple_period_rub / start_value * 100

            period_flows = [(start_dt, -start_value)]
            for r in operations_rows[1:]:
                op_type = str(r[H_ops['type']] or '')
                category = categorize_operation(op_type)
                try:
                    payment = float(r[H_ops['payment_rub']] or 0)
                except (ValueError, TypeError):
                    payment = 0.0
                dt = _parse_datetime(r[H_ops['date']])
                if dt is None or dt < start_dt: continue
                if category == 'input' and payment > 0:
                    period_flows.append((dt, -payment))
                elif category == 'output' and payment < 0:
                    period_flows.append((dt, -payment))
            period_flows.append((today, total_value))
            xirr_period = _xirr(period_flows)

    net_invested = total_inputs - total_outputs
    result_all_rub = total_value - net_invested

    year_passive = year_dividends + year_coupons + year_fees + year_taxes
    year_net_flow = year_inputs - year_outputs

    rows = [['Показатель', 'Значение', 'Комментарий']]

    rows.append(['📅 ПЕРИОД', '', ''])
    rows.append(['С', period_start.strftime('%d.%m.%Y'), f'скользящие {REPORT_WINDOW_DAYS} дней'])
    rows.append(['По', today.strftime('%d.%m.%Y'), 'сегодня'])
    rows.append(['', '', ''])

    rows.append(['💰 ДОХОДЫ ЗА ПЕРИОД', '', ''])
    rows.append(['Дивиденды', round(year_dividends, 2), ''])
    rows.append(['Купоны', round(year_coupons, 2), ''])
    rows.append(['💸 РАСХОДЫ ЗА ПЕРИОД', '', ''])
    rows.append(['Комиссии брокера', round(year_fees, 2), ''])
    rows.append(['Налоги', round(year_taxes, 2), ''])
    rows.append(['📊 ПАССИВНЫЙ ДОХОД', round(year_passive, 2),
                 'Дивиденды + купоны − комиссии − налоги'])
    rows.append(['', '', ''])

    rows.append(['💵 ДВИЖЕНИЕ ДЕНЕГ ЗА ПЕРИОД', '', ''])
    rows.append(['Пополнения', round(year_inputs, 2), ''])
    rows.append(['Выводы', round(year_outputs, 2), ''])
    rows.append(['Нетто-поток', round(year_net_flow, 2), 'Пополнения − выводы'])
    rows.append(['', '', ''])

    rows.append(['📈 ПОРТФЕЛЬ СЕЙЧАС', '', ''])
    rows.append(['Стоимость бумаг', round(portfolio_value, 2), 'без денег'])
    rows.append(['Свободные деньги', round(money_value, 2), ''])
    rows.append(['Всего', round(total_value, 2), 'бумаги + деньги'])
    rows.append(['', '', ''])

    rows.append([f'📈 РЕЗУЛЬТАТ ЗА {REPORT_WINDOW_DAYS} ДНЕЙ', '', ''])
    if start_value is not None:
        rows.append(['Стоимость на начало', round(start_value, 2),
                     f'снимок от {start_date_str}'])
        rows.append(['Стоимость сейчас', round(total_value, 2), ''])
        rows.append(['Пополнения за период', round(year_inputs, 2), ''])
        rows.append(['Выводы за период', round(year_outputs, 2), ''])
        rows.append(['Результат за период', round(simple_period_rub, 2),
                     'Сейчас − Начало − Пополнения + Выводы'])
    else:
        rows.append(['Результат за период', 'н/д',
                     f'накапливаем снимки (нужно {REPORT_WINDOW_DAYS} дней)'])
    rows.append(['', '', ''])

    rows.append([f'📋 ИТОГИ ЗА {OPERATIONS_DAYS} ДНЕЙ', '', ''])
    rows.append(['Внесено всего', round(total_inputs, 2), ''])
    rows.append(['Выведено всего', round(total_outputs, 2), ''])
    rows.append(['Вложено (нетто)', round(net_invested, 2), 'Внесено − выведено'])
    rows.append(['Результат в рублях', round(result_all_rub, 2),
                 'Портфель сейчас − вложено (нетто)'])
    rows.append(['   ├ дивиденды', round(total_dividends, 2), ''])
    rows.append(['   ├ купоны', round(total_coupons, 2), ''])
    rows.append(['   ├ комиссии', round(total_fees, 2), ''])
    rows.append(['   ├ налоги', round(total_taxes, 2), ''])
    rows.append(['   └ изменение цены',
                 round(result_all_rub - total_dividends - total_coupons - total_fees - total_taxes, 2),
                 'остаток'])
    rows.append(['', '', ''])

    rows.append(['📉 ДОХОДНОСТЬ', '', ''])
    if simple_period_pct is not None and actual_days:
        rows.append([f'Простая за {actual_days} дней, %',
                     round(simple_period_pct, 2), f'с {start_date_str}'])
    else:
        rows.append([f'Простая за {REPORT_WINDOW_DAYS} дней, %', 'н/д',
                     'накапливаем снимки'])

    if xirr_period is not None and actual_days:
        rows.append([f'XIRR за {actual_days} дней, %',
                     round(xirr_period, 2), 'с учётом дат потоков'])
    else:
        rows.append([f'XIRR за {REPORT_WINDOW_DAYS} дней, %', 'н/д',
                     'накапливаем снимки'])

    if xirr_all_time is not None:
        rows.append(['XIRR за всё время, %', round(xirr_all_time, 2),
                     'по всей истории (для справки)'])

    return rows


# ----------------------------------------------------------------------------
# Все листы
# ----------------------------------------------------------------------------
def build_all_sheets(raw_data, operations_data, snapshot_data, snapshot_365):
    header = raw_data[0]
    H = {k: i for i, k in enumerate(header)}
    sheets = {}

    sheets[RAW_SHEET] = raw_data
    if AGGREGATE_BY in ('name', 'ticker', 'figi'):
        sheets[AGG_SHEET] = aggregate_by_key(raw_data, AGGREGATE_BY)

    per_type = {cat: [header[:]] for cat in TYPE_SHEETS.keys()}
    for r in raw_data[1:]:
        is_money = str(r[H['type']]).lower() == 'money'
        cat = 'Money' if is_money else to_category(r[H['type']])
        if cat in per_type:
            per_type[cat].append(r)

    for cat, sheet_name in TYPE_SHEETS.items():
        sheets[sheet_name] = aggregate_by_key(per_type[cat], AGGREGATE_BY)

    summary = [['type', 'position_value_rub']]
    sums = {}
    for r in raw_data[1:]:
        cat = to_category(r[H['type']])
        val = float(r[H['position_value_rub']] or 0)
        sums[cat] = sums.get(cat, 0) + val
    for cat in ('Shares', 'Bonds', 'ETFs', 'Currencies', 'Futures', 'Other'):
        if cat in sums:
            summary.append([cat, sums[cat]])
    money_sum = sum(float(r[H['position_value_rub']] or 0) for r in per_type['Money'][1:])
    if money_sum > 0:
        summary.append(['Money', money_sum])
    sheets[SUMMARY_BY_TYPE_SHEET] = summary

    sheets[OPERATIONS_SHEET] = operations_data
    sheets[YEAR_REPORT_SHEET] = build_year_report(operations_data, raw_data, snapshot_365)
    sheets[SNAPSHOTS_SHEET] = snapshot_data

    return sheets


# ----------------------------------------------------------------------------
# Точка входа
# ----------------------------------------------------------------------------
def main():
    print("Подключение к Tinkoff Invest API...")
    with Client(TOKEN) as client:
        accounts_response = client.users.get_accounts()
        all_accounts = accounts_response.accounts
        accounts = [
            acc for acc in all_accounts
            if 'OPEN' in _status_to_str(getattr(acc, 'status', None))
            and 'CLOSED' not in _status_to_str(getattr(acc, 'status', None))
        ]
        print(f"Открытых счетов: {len(accounts)}")

        print("\nСбор позиций...")
        raw_data = collect_positions(client, accounts)
        print(f"Собрано строк: {len(raw_data) - 1}")

        operations_data = [['date', 'accountId', 'type', 'ticker', 'name',
                            'quantity', 'price_rub', 'payment_rub', 'currency',
                            'figi', 'operation_id']]
        if OPERATIONS_DAYS > 0:
            to_date = datetime.now(timezone.utc)
            from_date = to_date - timedelta(days=OPERATIONS_DAYS)
            print(f"\nСбор операций за период {from_date.date()} .. {to_date.date()}")
            operations_data = build_operations_data(client, accounts, from_date, to_date)
            print(f"\nСобрано операций: {len(operations_data) - 1}")

    client_api = AppsScriptClient(APPS_SCRIPT_URL, SECRET)

    print(f"\nЧтение снимка портфеля за {SNAPSHOT_DAYS_AGO} дней...")
    snapshot_365 = client_api.get_snapshot_n_days_ago(SNAPSHOT_DAYS_AGO)
    if snapshot_365:
        print(f"  Найден снимок от {snapshot_365['date']}: "
              f"{snapshot_365['total_value']:,.2f} ₽")
    else:
        print(f"  Снимок не найден — XIRR за скользящий период будет 'н/д'")

    snapshot_data = build_today_snapshot(raw_data)
    print(f"\nСегодняшний снимок: {snapshot_data[1][0]} — {snapshot_data[1][1]:,.2f} ₽")

    print("\nПодготовка листов...")
    sheets = build_all_sheets(raw_data, operations_data, snapshot_data, snapshot_365)

    print("Отправка в Google Sheets через Apps Script...")
    client_api.write_tables(sheets)

    print("Готово. Все листы обновлены.")


if __name__ == '__main__':
    main()