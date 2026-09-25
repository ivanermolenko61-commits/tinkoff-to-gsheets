#!/usr/bin/env bash
# Оповещение в Telegram об изменении состояния синка.
#
# Использование: notify.sh <success|failure|cancelled>
# Окружение: TG_TOKEN, TG_CHAT_ID (секреты GitHub), GH_TOKEN, GITHUB_REPOSITORY,
#            GITHUB_RUN_ID, GITHUB_SERVER_URL; TEST_NOTIFY=true — отправить тестовое.
#            PREV — для локальной проверки: «заранее известные» итоги прошлых запусков.
#
# Синк идёт каждые 15 минут, а одиночный сбой API Т-Банка — норма (следующий
# запуск обычно зелёный). Поэтому не шумим на каждый красный запуск:
#   • провал, и предыдущий тоже провал, а до него — нет  → «синк не работает» (1 раз);
#   • успех после ≥2 провалов подряд                     → «синк снова работает».
# Скрипт никогда не роняет job: иначе сам стал бы причиной «красного» запуска.

STATUS="${1:-}"
RUN_URL="${GITHUB_SERVER_URL:-https://github.com}/${GITHUB_REPOSITORY:-}/actions/runs/${GITHUB_RUN_ID:-}"

if [ -z "${TG_TOKEN:-}" ] || [ -z "${TG_CHAT_ID:-}" ]; then
  echo "Секреты TELEGRAM_BOT_TOKEN / TELEGRAM_CHAT_ID не заданы — оповещение пропущено"
  exit 0
fi

send() {
  code=$(curl -sS -o /dev/null -w "%{http_code}" -X POST \
    "https://api.telegram.org/bot${TG_TOKEN}/sendMessage" \
    --data-urlencode "chat_id=${TG_CHAT_ID}" \
    --data-urlencode "parse_mode=HTML" \
    --data-urlencode "disable_web_page_preview=true" \
    --data-urlencode "text=$1") || code="curl error"
  echo "Telegram: HTTP ${code}"
}

if [ "${TEST_NOTIFY:-false}" = "true" ]; then
  send "🧪 Тест оповещений синка Т-Инвестиций: всё настроено. Статус запуска: ${STATUS}.
<a href=\"${RUN_URL}\">Открыть запуск</a>"
  exit 0
fi

if [ "$STATUS" != "success" ] && [ "$STATUS" != "failure" ]; then
  echo "Статус '${STATUS}' — оповещение не нужно"
  exit 0
fi

# Итоги двух предыдущих ЗАВЕРШЁННЫХ запусков (только success/failure, новые — первыми).
# Отменённые (cancelled) пропускаем — их даёт concurrency, это не сбой.
if [ -z "${PREV+x}" ]; then
  PREV=$(gh api "repos/${GITHUB_REPOSITORY}/actions/workflows/sync.yml/runs?status=completed&per_page=20" \
    --jq "[.workflow_runs[] | select(.id != ${GITHUB_RUN_ID}) | select(.conclusion == \"success\" or .conclusion == \"failure\") | .conclusion][0:2] | join(\" \")" \
    2>/dev/null) || { echo "Не удалось получить историю запусков — оповещение пропущено"; exit 0; }
fi
set -- $PREV
P1="${1:-}"
P2="${2:-}"
echo "Текущий: ${STATUS}; предыдущие: ${P1:-нет} ${P2:-нет}"

if [ "$STATUS" = "failure" ] && [ "$P1" = "failure" ] && [ "$P2" != "failure" ]; then
  send "⚠️ <b>Синк Т-Инвестиций → Google Таблица не работает</b>
Два запуска подряд завершились ошибкой, данные в таблице не обновляются.

<a href=\"${RUN_URL}\">Открыть лог запуска</a>"
elif [ "$STATUS" = "success" ] && [ "$P1" = "failure" ] && [ "$P2" = "failure" ]; then
  send "✅ <b>Синк Т-Инвестиций снова работает</b>
Таблица обновлена.

<a href=\"${RUN_URL}\">Открыть запуск</a>"
else
  echo "Состояние не изменилось — без оповещения"
fi
exit 0
