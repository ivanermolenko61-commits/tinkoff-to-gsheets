"""
Tinkoff Invest → Google Sheets

Собирает позиции, операции и снимок портфеля по всем открытым счетам
и отправляет их в Apps Script, который раскладывает данные по листам.

Листы:
  Positions, Positions_Aggregated, Positions_Shares, Positions_Bonds,
  Positions_ETFs, Positions_Currencies, Positions_Futures, Positions_Other,
  Positions_Money, Positions_SummaryByType, Operations,
  Report (отчётная панель),
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
SNAPSHOT_MIN_DAYS_AGO = 2

RAW_SHEET = 'Positions'
AGG_SHEET = 'Positions_Aggregated'
SUMMARY_BY_TYPE_SHEET = 'Positions_SummaryByType'
OPERATIONS_SHEET = 'Operations'
REPORT_SHEET = 'Report'
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

    def get_oldest_snapshot(self, min_days_ago=2):
        try:
            resp = requests.get(
                self.url,
                params={
                    'secret': self.secret,
                    'action': 'get_oldest_snapshot',
                    'min_days_ago': min_days_ago,
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
    """Годовая доходность по потокам. cashflows: список (datetime, amount).
    Оттоки — с минусом, притоки — с плюсом."""
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
# Агрегация (без изменений)
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
# Сбор позиций (без изменений)
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
# Сбор операций (без изменений)
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
# Отчётная панель
# ----------------------------------------------------------------------------
def _fmt_money(v):
    return round(float(v or 0), 2)


def _pct(x, base):
    if base is None or base == 0:
        return None
    return round(x / base * 100, 2)


def build_report(operations_rows, raw_positions, snapshot_ref, today=None):
    """Формирует лист Report — человекочитаемую панель.

    Логика:
      Вложено (нетто)  = Пополнения − Выводы
      Результат (руб)  = Портфель сейчас − Вложено (нетто)
      Структура результата:
          Дивиденды + Купоны + Комиссии + Налоги + Изменение цены
      Простая доходность = Результат / Вложено (нетто) * 100
      Годовая (XIRR)     = по всем потокам капитала (input/output/final)
    """
    if today is None:
        today = datetime.now(timezone.utc).replace(tzinfo=None)

    H_ops = {k: i for i, k in enumerate(operations_rows[0])}
    H_pos = {k: i for i, k in enumerate(raw_positions[0])}

    # --- Портфель сейчас ---
    portfolio_value = 0.0
    money_value = 0.0
    for r in raw_positions[1:]:
        ptype = str(r[H_pos['type']] or '').lower()
        v = float(r[H_pos['position_value_rub']] or 0)
        if ptype == 'money':
            money_value += v
        else:
            portfolio_value += v
    total_value = portfolio_value + money_value

    # --- Все потоки и доходы за всё время ---
    total_inputs = 0.0
    total_outputs = 0.0
    total_dividends = 0.0
    total_coupons = 0.0
    total_fees = 0.0          # знак отрицательный, как в операциях
    total_taxes = 0.0         # знак отрицательный
    xirr_flows_all = []
    first_op_date = None
    last_op_date = None

    for r in operations_rows[1:]:
        op_type = str(r[H_ops['type']] or '')
        category = categorize_operation(op_type)
        try:
            payment = float(r[H_ops['payment_rub']] or 0)
        except (ValueError, TypeError):
            payment = 0.0
        dt = _parse_datetime(r[H_ops['date']])
        if dt is not None:
            if first_op_date is None or dt < first_op_date: first_op_date = dt
            if last_op_date is None or dt > last_op_date: last_op_date = dt

        if category == 'input' and payment > 0:
            total_inputs += payment
            if dt is not None:
                xirr_flows_all.append((dt, -payment))
        elif category == 'output' and payment < 0:
            total_outputs += abs(payment)
            if dt is not None:
                xirr_flows_all.append((dt, abs(payment)))
        elif category == 'dividend':
            total_dividends += payment
        elif category == 'coupon':
            total_coupons += payment
        elif category == 'fee':
            total_fees += payment
        elif category == 'tax':
            total_taxes += payment

    net_invested = total_inputs - total_outputs
    result_rub = total_value - net_invested
    # Изменение цены = всё, что не объяснено доходами и расходами
    price_change = result_rub - total_dividends - total_coupons - total_fees - total_taxes

    # XIRR за всё время
    xirr_all = None
    if xirr_flows_all:
        xirr_flows_all.append((today, total_value))
        xirr_all = _xirr(xirr_flows_all)

    simple_all_pct = _pct(result_rub, net_invested) if net_invested > 0 else None
    portfolio_age_days = (today - first_op_date).days if first_op_date else None

    # --- Период от снимка ---
    period_days = None
    period_start_str = None
    period_simple_pct = None
    period_xirr = None
    period_result_rub = None
    period_inputs = 0.0
    period_outputs = 0.0

    if snapshot_ref and snapshot_ref.get('date'):
        start_dt = _parse_datetime(snapshot_ref['date'])
        if start_dt:
            start_value = float(snapshot_ref.get('total_value') or 0)
            period_days = (today - start_dt).days
            period_start_str = snapshot_ref['date']

            for r in operations_rows[1:]:
                op_type = str(r[H_ops['type']] or '')
                category = categorize_operation(op_type)
                try:
                    payment = float(r[H_ops['payment_rub']] or 0)
                except (ValueError, TypeError):
                    payment = 0.0
                dt = _parse_datetime(r[H_ops['date']])
                if dt is None or dt < start_dt:
                    continue
                if category == 'input' and payment > 0:
                    period_inputs += payment
                elif category == 'output' and payment < 0:
                    period_outputs += abs(payment)

            period_result_rub = total_value - start_value - period_inputs + period_outputs
            period_simple_pct = _pct(period_result_rub, start_value) if start_value > 0 else None

            period_flows = [(start_dt, -start_value)]
            for r in operations_rows[1:]:
                op_type = str(r[H_ops['type']] or '')
                category = categorize_operation(op_type)
                try:
                    payment = float(r[H_ops['payment_rub']] or 0)
                except (ValueError, TypeError):
                    payment = 0.0
                dt = _parse_datetime(r[H_ops['date']])
                if dt is None or dt < start_dt:
                    continue
                if category == 'input' and payment > 0:
                    period_flows.append((dt, -payment))
                elif category == 'output' and payment < 0:
                    period_flows.append((dt, abs(payment)))
            period_flows.append((today, total_value))
            period_xirr = _xirr(period_flows)

    # --- Сборка листа ---
    rows = []
    rows.append(['📊 ОТЧЁТ ПО ПОРТФЕЛЮ', '', ''])
    rows.append([f'на {today.strftime("%d.%m.%Y")}', '', ''])
    if portfolio_age_days is not None:
        rows.append([f'история операций: {portfolio_age_days} дней', '', ''])
    rows.append(['', '', ''])

    # --- Портфель сейчас ---
    rows.append(['💰 ПОРТФЕЛЬ СЕЙЧАС', 'Сумма, ₽', 'Доля'])
    rows.append(['Бумаги', _fmt_money(portfolio_value),
                 f'{portfolio_value/total_value*100:.1f}%' if total_value else '—'])
    rows.append(['Свободные деньги', _fmt_money(money_value),
                 f'{money_value/total_value*100:.1f}%' if total_value else '—'])
    rows.append(['ИТОГО', _fmt_money(total_value), '100.0%'])
    rows.append(['', '', ''])

    # --- Вложения ---
    rows.append(['💵 ВЛОЖЕНИЯ ЗА ВСЁ ВРЕМЯ', 'Сумма, ₽', ''])
    rows.append(['Пополнения', _fmt_money(total_inputs), ''])
    rows.append(['Выводы', _fmt_money(total_outputs), ''])
    rows.append(['Вложено (нетто)', _fmt_money(net_invested), 'Пополнения − выводы'])
    rows.append(['', '', ''])

    # --- Результат ---
    rows.append(['📈 РЕЗУЛЬТАТ ЗА ВСЁ ВРЕМЯ', 'Сумма, ₽', '% от вложенного'])
    result_pct = _pct(result_rub, net_invested) if net_invested > 0 else None
    rows.append(['Портфель − вложено', _fmt_money(result_rub),
                 f'{result_pct:.2f}%' if result_pct is not None else '—'])
    rows.append(['', '', ''])
    rows.append(['  Структура результата:', '', ''])

    def _pct_of_invested(x):
        p = _pct(x, net_invested) if net_invested > 0 else None
        return f'{p:.2f}%' if p is not None else '—'

    rows.append(['  ├ Дивиденды', _fmt_money(total_dividends), _pct_of_invested(total_dividends)])
    rows.append(['  ├ Купоны', _fmt_money(total_coupons), _pct_of_invested(total_coupons)])
    rows.append(['  ├ Комиссии', _fmt_money(total_fees), _pct_of_invested(total_fees)])
    rows.append(['  ├ Налоги', _fmt_money(total_taxes), _pct_of_invested(total_taxes)])
    rows.append(['  └ Изменение цены', _fmt_money(price_change), _pct_of_invested(price_change)])
    rows.append(['', '', ''])

    # --- Доходность за всё время ---
    rows.append(['📉 ДОХОДНОСТЬ ЗА ВСЁ ВРЕМЯ', '%', ''])
    if simple_all_pct is not None:
        rows.append(['Простая', f'{simple_all_pct:.2f}%',
                     f'за {portfolio_age_days} дн.' if portfolio_age_days else ''])
    else:
        rows.append(['Простая', 'н/д', 'недостаточно данных'])
    if xirr_all is not None:
        rows.append(['Годовая (XIRR)', f'{xirr_all:.2f}%', 'с учётом дат потоков'])
    else:
        rows.append(['Годовая (XIRR)', 'н/д', 'нужны пополнения и выводы'])
    rows.append(['', '', ''])

    # --- Доходность за период от снимка ---
    period_label = period_days if period_days is not None else REPORT_WINDOW_DAYS
    rows.append([f'📉 ДОХОДНОСТЬ ЗА {period_label} ДНЕЙ (от снимка)', '%', ''])
    if period_simple_pct is not None:
        rows.append(['Простая', f'{period_simple_pct:.2f}%',
                     f'с {period_start_str}'])
    else:
        rows.append(['Простая', 'н/д',
                     f'накапливаем снимки (нужно {REPORT_WINDOW_DAYS} дней)'])
    if period_xirr is not None:
        rows.append(['XIRR за период', f'{period_xirr:.2f}%', 'годовая, с учётом дат потоков'])
    else:
        rows.append(['XIRR за период', 'н/д',
                     f'накапливаем снимки (нужно {REPORT_WINDOW_DAYS} дней)'])

    return rows


# ----------------------------------------------------------------------------
# Сборка всех листов
# ----------------------------------------------------------------------------
def build_all_sheets(raw_data, operations_data, snapshot_data, snapshot_ref):
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
    sheets[REPORT_SHEET] = build_report(operations_data, raw_data, snapshot_ref)
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
    snapshot_ref = client_api.get_snapshot_n_days_ago(SNAPSHOT_DAYS_AGO)
    if snapshot_ref:
        print(f"  Найден снимок от {snapshot_ref['date']}: "
              f"{snapshot_ref['total_value']:,.2f} ₽")
    else:
        print(f"  Снимка за {SNAPSHOT_DAYS_AGO} дней нет — беру самый старый "
              f"(не моложе {SNAPSHOT_MIN_DAYS_AGO} дней)...")
        snapshot_ref = client_api.get_oldest_snapshot(min_days_ago=SNAPSHOT_MIN_DAYS_AGO)
        if snapshot_ref:
            print(f"  Использую снимок от {snapshot_ref['date']}: "
                  f"{snapshot_ref['total_value']:,.2f} ₽")
        else:
            print("  Снимков нужной давности нет — доходность за период будет 'н/д'")

    snapshot_data = build_today_snapshot(raw_data)
    print(f"\nСегодняшний снимок: {snapshot_data[1][0]} — {snapshot_data[1][1]:,.2f} ₽")

    print("\nПодготовка листов...")
    sheets = build_all_sheets(raw_data, operations_data, snapshot_data, snapshot_ref)

    print("Отправка в Google Sheets через Apps Script...")
    client_api.write_tables(sheets)

    print("Готово. Все листы обновлены.")


if __name__ == '__main__':
    main()