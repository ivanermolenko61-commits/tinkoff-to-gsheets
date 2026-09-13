/**
 * Приём данных от Python-скрипта и запись их на листы Google Таблицы.
 *
 * Развёртывание:
 *   Развернуть → Управление развертываниями → Новая версия
 *     Запуск от имени: Я
 *     Кто имеет доступ: Все
 *
 * В Настройках проекта → Свойства скрипта должно быть:
 *   APPS_SECRET = <тот же секрет, что в .env Python-скрипта>
 */

const SECRET = PropertiesService.getScriptProperties().getProperty('APPS_SECRET') || '';

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
      return jsonResponse_({
        status: 'error',
        message: 'Payload не получен',
        debug: debugInfo,
      });
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

    for (var sheetName in data.sheets) {
      var rows = data.sheets[sheetName];
      result.push(sheetName + ': ' + (rows ? rows.length : 'NULL') + ' строк');

      if (!rows || rows.length === 0) continue;

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

      var cleanRows = rows.map(function(r) {
        var out = r.slice(0, width);
        while (out.length < width) out.push('');
        return out.map(function(v) {
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

        ['quantity_lots', 'quantity_pcs'].forEach(function(k) {
          if (idx[k]) sheet.getRange(2, idx[k], n, 1).setNumberFormat('0.########');
        });
        ['current_price_rub_per_piece', 'avg_price_rub_per_piece',
         'price_rub_per_lot', 'position_value_rub'].forEach(function(k) {
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