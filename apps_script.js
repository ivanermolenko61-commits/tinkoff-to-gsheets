/**
 * Приём данных от Python-скрипта и запись их на листы Google Таблицы.
 *
 * ВАЖНО: после правки обязательно
 *   Развернуть → Управление развёртываниями → (карандаш) → Версия: "Новая версия" → Развернуть
 * Иначе Web App продолжит работать по старому коду.
 */

const SECRET = PropertiesService.getScriptProperties().getProperty('APPS_SECRET') || '';
const SCRIPT_VERSION = 'v5-2026-09-20';

// ---------------------------------------------------------------------------
// POST — запись листов
// ---------------------------------------------------------------------------
function doPost(e) {
  try {
    var payloadStr = null;
    var debugInfo = {
      hasPostData: !!e.postData,
      postDataType: e.postData ? e.postData.type : '',
      postDataLength: e.postData && e.postData.contents ? e.postData.contents.length : 0,
    };

    if (e.parameter && e.parameter.payload) {
      payloadStr = e.parameter.payload;
      debugInfo.source = 'e.parameter.payload';
    } else if (e.postData && e.postData.contents) {
      var body = e.postData.contents;
      if (body.indexOf('payload=') !== -1) {
        var params = parseFormUrlEncoded_(body);
        payloadStr = params.payload;
        debugInfo.source = 'form-urlencoded';
      } else if (body.charAt(0) === '{') {
        payloadStr = body;
        debugInfo.source = 'json-body';
      } else {
        var params2 = parseFormUrlEncoded_(body);
        payloadStr = params2.payload || null;
        debugInfo.source = 'form-urlencoded-fallback';
      }
    }

    if (!payloadStr) {
      return jsonResponse_({ status: 'error', message: 'Payload не получен', debug: debugInfo });
    }

    var data;
    try {
      data = JSON.parse(payloadStr);
    } catch (parseErr) {
      return jsonResponse_({
        status: 'error',
        message: 'JSON parse error: ' + parseErr.toString(),
        preview: payloadStr.substring(0, 200),
        debug: debugInfo,
      });
    }

    var received = data.secret === null || data.secret === undefined ? '' : String(data.secret);
    if (received.trim() !== String(SECRET).trim()) {
      return jsonResponse_({
        status: 'error',
        message: 'Invalid secret',
        debug: Object.assign(debugInfo, {
          receivedSecretLength: received.length,
          expectedSecretLength: String(SECRET).length,
        }),
      });
    }

    var ss = SpreadsheetApp.getActiveSpreadsheet();
    var result = [];
    result.push('Spreadsheet: ' + (ss ? ss.getName() : 'NULL'));
    result.push('Script version: ' + SCRIPT_VERSION);

    for (var sheetName in data.sheets) {
      var rows = data.sheets[sheetName];
      result.push(sheetName + ': ' + (rows ? rows.length : 'NULL') + ' строк');

      if (!rows || rows.length === 0) continue;

      // --- Накопление снимков: не перезаписываем, а append/update по дате ---
      if (sheetName === 'PortfolioSnapshots') {
        appendOrUpdateSnapshots_(ss, rows, result);
        continue;
      }
      // -----------------------------------------------------------------------

      var sheet = ss.getSheetByName(sheetName);
      if (!sheet) {
        sheet = ss.insertSheet(sheetName);
        result.push('  создан лист ' + sheetName);
      } else {
        var f = sheet.getFilter();
        if (f) f.remove();
        sheet.clearContents();
      }

      var width = rows[0].length;
      if (width === 0) {
        result.push('  пропущен (0 колонок)');
        continue;
      }

      var cleanRows = rows.map(function (r) {
        var out = r.slice(0, width);
        while (out.length < width) out.push('');
        return out.map(function (v) {
          if (v === null || v === undefined) return '';
          return v;
        });
      });

      sheet.getRange(1, 1, cleanRows.length, width).setValues(cleanRows);
      sheet.getRange(1, 1, 1, width).setFontWeight('bold');
      sheet.setFrozenRows(1);
      sheet.autoResizeColumns(1, width);

      var n = cleanRows.length - 1;
      if (n > 0) {
        var header = cleanRows[0];
        var idx = {};
        for (var i = 0; i < header.length; i++) idx[header[i]] = i + 1;

        ['quantity_lots', 'quantity_pcs'].forEach(function (k) {
          if (idx[k]) sheet.getRange(2, idx[k], n, 1).setNumberFormat('0.########');
        });
        ['current_price_rub_per_piece', 'avg_price_rub_per_piece', 'price_rub_per_lot', 'position_value_rub'].forEach(function (k) {
          if (idx[k]) sheet.getRange(2, idx[k], n, 1).setNumberFormat('#,##0.00 [$₽-ru-RU]');
        });

        sheet.getRange(1, 1, cleanRows.length, width).createFilter();
      }
    }

    return jsonResponse_({ status: 'ok', debug: result, source: debugInfo.source });
  } catch (err) {
    return jsonResponse_({ status: 'error', message: err.toString() });
  }
}

// ---------------------------------------------------------------------------
// GET — чтение снимков (для расчёта XIRR за скользящий период)
// ---------------------------------------------------------------------------
function doGet(e) {
  try {
    var secret = (e && e.parameter && e.parameter.secret) || '';
    if (String(secret).trim() !== String(SECRET).trim()) {
      return jsonResponse_({ status: 'error', message: 'Invalid secret' });
    }
    var action = ((e.parameter.action || '') + '').trim();

    if (action === 'ping') {
      return jsonResponse_({ status: 'ok', version: SCRIPT_VERSION });
    }

    if (action === 'get_snapshot_n_days_ago') {
      var daysAgo = parseInt(e.parameter.days_ago, 10);
      if (isNaN(daysAgo) || daysAgo < 0) daysAgo = 365;
      var snapshot = getSnapshotNDaysAgo_(daysAgo);
      return jsonResponse_({ status: 'ok', snapshot: snapshot });
    }

    if (action === 'get_oldest_snapshot') {
      var minDays = parseInt(e.parameter.min_days_ago, 10);
      if (isNaN(minDays) || minDays < 0) minDays = 2;
      var snap = getOldestSnapshot_(minDays);
      return jsonResponse_({ status: 'ok', snapshot: snap });
    }

    return jsonResponse_({ status: 'error', message: 'Unknown action: ' + action });
  } catch (err) {
    return jsonResponse_({ status: 'error', message: err.toString() });
  }
}

function getSnapshotNDaysAgo_(daysAgo) {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var sheet = ss.getSheetByName('PortfolioSnapshots');
  if (!sheet) return null;
  var data = sheet.getDataRange().getValues();
  if (data.length < 2) return null;

  var header = data[0].map(function (v) { return String(v); });
  var dateIdx = header.indexOf('date');
  var totalIdx = header.indexOf('total_value');
  var portIdx = header.indexOf('portfolio_value');
  var moneyIdx = header.indexOf('money_value');
  if (dateIdx === -1 || totalIdx === -1) return null;

  var now = new Date();
  var target = new Date(now.getTime() - daysAgo * 86400000);
  var targetStr = target.getFullYear() + '-' + pad2_(target.getMonth() + 1) + '-' + pad2_(target.getDate());

  var best = null;
  var bestDiff = Infinity;
  for (var i = 1; i < data.length; i++) {
    var d = normalizeDate_(data[i][dateIdx]);
    if (!d) continue;
    var diff = Math.abs(daysBetween_(d, targetStr));
    if (diff < bestDiff) {
      bestDiff = diff;
      best = {
        date: d,
        total_value: Number(data[i][totalIdx]) || 0,
        portfolio_value: portIdx >= 0 ? Number(data[i][portIdx]) || 0 : 0,
        money_value: moneyIdx >= 0 ? Number(data[i][moneyIdx]) || 0 : 0,
      };
    }
  }
  // Если ближайший снимок дальше, чем на 60 дней от целевой даты — считаем, что снимка нет
  if (best && bestDiff > 60) return null;
  return best;
}

/**
 * Возвращает самый старый снимок, которому не меньше minDaysAgo дней.
 * Нужно для случая, когда снимка за 365 дней ещё нет, но накопление уже идёт:
 * тогда доходность считается от самого старого доступного снимка.
 */
function getOldestSnapshot_(minDaysAgo) {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var sheet = ss.getSheetByName('PortfolioSnapshots');
  if (!sheet) return null;
  var data = sheet.getDataRange().getValues();
  if (data.length < 2) return null;

  var header = data[0].map(function (v) { return String(v); });
  var dateIdx = header.indexOf('date');
  var totalIdx = header.indexOf('total_value');
  var portIdx = header.indexOf('portfolio_value');
  var moneyIdx = header.indexOf('money_value');
  if (dateIdx === -1 || totalIdx === -1) return null;

  var now = new Date();
  var todayStr = now.getFullYear() + '-' + pad2_(now.getMonth() + 1) + '-' + pad2_(now.getDate());

  var best = null;
  var bestAge = -1;
  for (var i = 1; i < data.length; i++) {
    var d = normalizeDate_(data[i][dateIdx]);
    if (!d) continue;
    var age = daysBetween_(todayStr, d);
    if (age < minDaysAgo) continue;
    if (age > bestAge) {
      bestAge = age;
      best = {
        date: d,
        total_value: Number(data[i][totalIdx]) || 0,
        portfolio_value: portIdx >= 0 ? Number(data[i][portIdx]) || 0 : 0,
        money_value: moneyIdx >= 0 ? Number(data[i][moneyIdx]) || 0 : 0,
      };
    }
  }
  return best;
}

// ---------------------------------------------------------------------------
// Накопление снимков
// ---------------------------------------------------------------------------
function appendOrUpdateSnapshots_(ss, rows, result) {
  var sheetName = 'PortfolioSnapshots';
  var sheet = ss.getSheetByName(sheetName);
  if (!sheet) {
    sheet = ss.insertSheet(sheetName);
    result.push('  создан лист ' + sheetName);
  }

  var f = sheet.getFilter();
  if (f) f.remove();

  var header = rows[0];
  var dataRows = rows.slice(1);
  if (dataRows.length === 0) {
    result.push('  нет данных для снимков');
    return;
  }

  var width = header.length;
  var cleanHeader = header.map(function (v) {
    return v === null || v === undefined ? '' : String(v);
  });

  var existing = sheet.getDataRange().getValues();
  var headerMatches = existing.length > 0 &&
    existing[0].map(function (v) { return String(v); }).join('|') === cleanHeader.join('|');

  if (!headerMatches) {
    sheet.clearContents();
    var cleanData = dataRows.map(function (r) {
      var out = r.slice(0, width);
      while (out.length < width) out.push('');
      return out;
    });
    var all = [cleanHeader].concat(cleanData);
    sheet.getRange(1, 1, all.length, width).setValues(all);
    sheet.getRange(1, 1, 1, width).setFontWeight('bold');
    sheet.setFrozenRows(1);
    sheet.autoResizeColumns(1, width);
    result.push('  записана шапка и ' + dataRows.length + ' строк (было пусто)');
    return;
  }

  var dateCol = -1;
  for (var i = 0; i < cleanHeader.length; i++) {
    if (cleanHeader[i].toLowerCase() === 'date') { dateCol = i; break; }
  }

  var existingDates = {};
  if (dateCol >= 0) {
    for (var r = 1; r < existing.length; r++) {
      var k = normalizeDate_(existing[r][dateCol]);
      if (k) existingDates[k] = r + 1;
    }
  }

  var added = 0, updated = 0;
  for (var j = 0; j < dataRows.length; j++) {
    var row = dataRows[j].slice(0, width);
    while (row.length < width) row.push('');
    var key = dateCol >= 0 ? normalizeDate_(row[dateCol]) : '';

    if (key && existingDates[key]) {
      sheet.getRange(existingDates[key], 1, 1, width).setValues([row]);
      updated++;
    } else {
      sheet.appendRow(row);
      if (key) existingDates[key] = sheet.getLastRow();
      added++;
    }
  }
  result.push('  добавлено: ' + added + ', обновлено: ' + updated);
}

// ---------------------------------------------------------------------------
// Хелперы
// ---------------------------------------------------------------------------
function normalizeDate_(v) {
  if (v === null || v === undefined || v === '') return '';
  if (v instanceof Date) {
    return v.getFullYear() + '-' + pad2_(v.getMonth() + 1) + '-' + pad2_(v.getDate());
  }
  var s = String(v).trim();
  var m = s.match(/^(\d{4})-(\d{2})-(\d{2})/);
  if (m) return m[1] + '-' + m[2] + '-' + m[3];
  var parsed = new Date(s);
  if (!isNaN(parsed.getTime())) {
    return parsed.getFullYear() + '-' + pad2_(parsed.getMonth() + 1) + '-' + pad2_(parsed.getDate());
  }
  return s;
}

function daysBetween_(a, b) {
  var pa = a.split('-').map(Number);
  var pb = b.split('-').map(Number);
  var da = Date.UTC(pa[0], pa[1] - 1, pa[2]);
  var db = Date.UTC(pb[0], pb[1] - 1, pb[2]);
  return Math.round((da - db) / 86400000);
}

function pad2_(n) { return n < 10 ? '0' + n : '' + n; }

function parseFormUrlEncoded_(body) {
  var params = {};
  if (!body) return params;
  var pairs = body.split('&');
  for (var i = 0; i < pairs.length; i++) {
    var pair = pairs[i];
    var eq = pair.indexOf('=');
    if (eq === -1) continue;
    var k = decodeURIComponent(pair.substring(0, eq).replace(/\+/g, ' '));
    var v = decodeURIComponent(pair.substring(eq + 1).replace(/\+/g, ' '));
    params[k] = v;
  }
  return params;
}

function jsonResponse_(obj) {
  return ContentService
    .createTextOutput(JSON.stringify(obj))
    .setMimeType(ContentService.MimeType.JSON);
}
