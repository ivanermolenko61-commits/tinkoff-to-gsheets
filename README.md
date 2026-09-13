# Tinkoff Invest → Google Sheets

Синхронизация портфеля из Tinkoff Invest API в Google Таблицу через Apps Script.

## Зачем нужен такой подход

Google Apps Script не может напрямую обращаться к Tinkoff Invest API: сертификат
Минцифры РФ не входит в доверенные в среде Apps Script, а параметр
`validateHttpsCertificates: false` эту проблему не решает. Поэтому используется
двухслойная схема:

1. **Python-скрипт** локально получает данные через официальный SDK
   [`t-tech-investments`](https://opensource.tbank.ru/invest/), в котором сертификат
   уже встроен.
2. Скрипт отправляет данные POST-запросом в **Apps Script**.
3. Apps Script раскладывает данные по листам Google Таблицы.

## Что создаётся в таблице

| Лист | Что содержит |
|---|---|
| `Positions` | Сырые позиции по всем счетам |
| `Positions_Aggregated` | Агрегировано по `name` / `ticker` / `figi` |
| `Positions_Shares` | Только акции, агрегировано |
| `Positions_Bonds` | Только облигации, агрегировано |
| `Positions_ETFs` | Только фонды, агрегировано |
| `Positions_Currencies` | Только валюты, агрегировано |
| `Positions_Futures` | Только фьючерсы, агрегировано |
| `Positions_Other` | Всё остальное, агрегировано |
| `Positions_Money` | Деньги (RUB) по всем счетам |
| `Positions_SummaryByType` | Свод сумм в ₽ по категориям |

## Установка

### Требования

- Python 3.11+
- Google-аккаунт (для таблицы)

### 1. Клонировать репозиторий

```bash
git clone https://github.com/your-username/tinkoff-to-gsheets.git
cd tinkoff-to-gsheets