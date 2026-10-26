/**
 * @OnlyCurrentDoc
 */

/**
 * WebアプリケーションとしてアクセスされたときにHTML画面を表示する関数
 */
function doGet() {
  // Index.html（またはお使いのHTMLファイル名）を読み込んで画面を構築します
  return HtmlService.createTemplateFromFile('index')
      .evaluate()
      .setTitle('在庫・棚卸管理システム')
      .addMetaTag('viewport', 'width=device-width, initial-scale=1.0')
      .setXFrameOptionsMode(HtmlService.XFrameOptionsMode.ALLOWALL);
}

/**
 * 全データの取得
 */
function getAllData() {
  var ss = SpreadsheetApp.getActiveSpreadsheet();

  return {
    products: getSheetDataAsObjects(ss, "products"),
    history: getSheetDataAsObjects(ss, "history"),
    lockedMonths: getSheetDataAsObjects(ss, "lockedMonths").map(function(item) { 
      return item.month || item.value || item; 
    }),
    supplyItems: getSheetDataAsObjects(ss, "supplyItems").map(function(item) {
      return { name: item.name || item.value || "", note: item.note || "" };
    })
  };
}

var ALLOWED_SHEET_NAMES = {
  products: true,
  history: true,
  lockedMonths: true,
  supplyItems: true
};

function validateSheetName_(sheetName) {
  if (typeof sheetName !== 'string' || !ALLOWED_SHEET_NAMES[sheetName]) {
    throw new Error('許可されていないシート名です。');
  }
}

function validateScalarValue_(value, fieldName) {
  if (value === null || value === undefined || value === '') return;
  if (typeof value === 'number' && !isFinite(value)) {
    throw new Error(fieldName + 'に不正な数値が含まれています。');
  }
  if (typeof value !== 'string' && typeof value !== 'number' && typeof value !== 'boolean') {
    throw new Error(fieldName + 'に不正な値が含まれています。');
  }
}

function validateSheetData_(sheetName, dataArray) {
  validateSheetName_(sheetName);
  if (!Array.isArray(dataArray)) {
    throw new Error('保存データが配列ではありません。');
  }

  dataArray.forEach(function(item, index) {
    if (item === null || item === undefined) {
      throw new Error('保存データの' + (index + 1) + '行目が不正です。');
    }
    if (typeof item === 'object' && !(item instanceof Date)) {
      Object.keys(item).forEach(function(key) {
        validateScalarValue_(item[key], '保存データの' + (index + 1) + '行目');
      });
    } else {
      validateScalarValue_(item, '保存データの' + (index + 1) + '行目');
    }
  });
}

function validateHistoryRow_(cleanRow) {
  if (!cleanRow || typeof cleanRow !== 'object' || Array.isArray(cleanRow)) {
    throw new Error('履歴データが不正です。');
  }

  ['id', 'date', 'isoDate', 'purchaseDate', 'product', 'productCategory', 'type', 'qty', 'price', 'recordedAt']
    .forEach(function(fieldName) {
      validateScalarValue_(cleanRow[fieldName], '履歴データの' + fieldName);
    });

  if (cleanRow.id === null || cleanRow.id === undefined || cleanRow.id === '') {
    throw new Error('履歴データのIDがありません。');
  }
}

function acquireDocumentLock_() {
  var lock = LockService.getDocumentLock();
  if (!lock.tryLock(10000)) {
    throw new Error('別の保存処理が実行中です。少し待ってから再実行してください。');
  }
  return lock;
}

/**
 * 旧バージョンとの互換性のために残す日次トリガーの呼び出し先。
 *
 * 履歴は削除しない。60日を過ぎた履歴を隠すかどうかは、クライアント画面側で
 * 切り替える。既存の時間主導型トリガーが存在しても、この関数で履歴が失われる
 * ことはない。
 */
function pruneOldHistoryRows() {
  return 0;
}

/**
 * pruneOldHistoryRows() を毎日実行するための旧トリガーを設定する。
 *
 * ★★★ 重要：この関数は Apps Script エディタの画面から「一度だけ手動で実行」してください。★★★
 *   （doGet や getAllData からは呼び出されません。トリガーは一度設定すれば
 *   スプレッドシート／プロジェクトに保存され、以後は毎日深夜3時ごろに自動的に
 *   pruneOldHistoryRows が実行され続けます。再度実行しても重複登録されないよう、
 *   既存の同名トリガーは事前に削除してから登録し直します。）
 *
 * 手順：Apps Scriptエディタでこのファイルを開く → 関数選択で
 *       「setupHistoryPruneTrigger」を選択 → ▶実行ボタンを押す → 権限を承認する。
 *
 * 現在の pruneOldHistoryRows() は履歴を削除しないため、このトリガーの設定は任意。
 */
function setupHistoryPruneTrigger() {
  var triggers = ScriptApp.getProjectTriggers();
  triggers.forEach(function(t) {
    if (t.getHandlerFunction() === 'pruneOldHistoryRows') {
      ScriptApp.deleteTrigger(t);
    }
  });

  ScriptApp.newTrigger('pruneOldHistoryRows')
    .timeBased()
    .everyDays(1)
    .atHour(3)
    .create();
}

/**
 * シートデータを安全にオブジェクト配列に変換する共通関数
 */
function getSheetDataAsObjects(ss, sheetName) {
  validateSheetName_(sheetName);
  var sheet = ss.getSheetByName(sheetName);
  if (!sheet) return [];
  
  var values = sheet.getDataRange().getValues();
  if (values.length <= 1) return [];
  
  var headers = values[0];
  var result = [];
  
  for (var i = 1; i < values.length; i++) {
    var row = values[i];
    var obj = {};
    var hasData = false;
    
    for (var j = 0; j < headers.length; j++) {
      var header = headers[j];
      if (!header) continue;
      
      var val = row[j];
      
      // 日付オブジェクト（Date）が送られてきた場合の型安全処理
      if (val instanceof Date) {
        if (header === "isoDate" || header === "recordedAt") {
          val = val.toISOString();
        } else if (header === "month") {
          // lockedMonths シートの month 列（"YYYY-MM"）が誤って日付型セルに
          // 変換されてしまっていた場合の自己修復処理。
          var my = val.getFullYear();
          var mm2 = ("0" + (val.getMonth() + 1)).slice(-2);
          val = my + "-" + mm2;
        } else {
          // 通常の日付表示（YYYY-MM-DD HH:mm）に安全変換
          var y = val.getFullYear();
          var m = ("0" + (val.getMonth() + 1)).slice(-2);
          var d = ("0" + val.getDate()).slice(-2);
          var hh = ("0" + val.getHours()).slice(-2);
          var mm = ("0" + val.getMinutes()).slice(-2);
          val = y + "-" + m + "-" + d + " " + hh + ":" + mm;
        }
      }
      
      obj[header] = val;
      if (val !== "" && val !== null && val !== undefined) {
        hasData = true;
      }
    }
    if (hasData) {
      result.push(obj);
    }
  }
  return result;
}

/**
 * シートへのデータ一括安全保存
 * 勝手な日付自動生成や型の破壊を物理的に防ぎます
 */
function saveSheetData(sheetName, dataArray) {
  validateSheetData_(sheetName, dataArray);
  var lock = acquireDocumentLock_();
  try {
    var ss = SpreadsheetApp.getActiveSpreadsheet();
    var sheet = ss.getSheetByName(sheetName);
  
  if (!sheet) {
    sheet = ss.insertSheet(sheetName);
  }
  
  sheet.clearContents();
  
  if (!dataArray || dataArray.length === 0) {
    // 最低限のヘッダーのみ再構築
    if (sheetName === "products") sheet.appendRow(["name", "category", "defaultPrice", "status", "createdMonth", "reducedTax", "inactiveMonth", "reactivateMonth", "checkSheetVisible", "note"]);
    if (sheetName === "history") sheet.appendRow(["id", "date", "isoDate", "purchaseDate", "product", "productCategory", "type", "qty", "price", "recordedAt"]);
    if (sheetName === "lockedMonths") sheet.appendRow(["month"]);
    if (sheetName === "supplyItems") sheet.appendRow(["name", "note"]);
    return;
  }
  
  // オブジェクトのキーからヘッダーを動的に抽出
  var headers = [];
  if (sheetName === "products") headers = ["name", "category", "defaultPrice", "status", "createdMonth", "reducedTax", "inactiveMonth", "reactivateMonth", "checkSheetVisible", "note"];
  else if (sheetName === "history") headers = ["id", "date", "isoDate", "purchaseDate", "product", "productCategory", "type", "qty", "price", "recordedAt"];
  else if (sheetName === "lockedMonths") headers = ["month"];
  else if (sheetName === "supplyItems") headers = ["name", "note"];
  else {
    // 予備ロジック
    for (var key in dataArray[0]) {
      if (headers.indexOf(key) === -1) headers.push(key);
    }
  }
  
  var outputRows = [headers];
  
  for (var i = 0; i < dataArray.length; i++) {
    var item = dataArray[i];
    var row = [];
    
    // 単純な文字列/数値配列だった場合のフォールバック（lockedMonths用など）
    if (typeof item !== 'object' || item === null) {
      row.push(item);
    } else {
      for (var j = 0; j < headers.length; j++) {
        var header = headers[j];
        var val = item[header];
        
        // 保存時に不必要・不正な型に変換されるのを防ぐガード
        if (val === undefined || val === null) {
          val = "";
        } else if (val instanceof Date) {
          val = val.toISOString();
        }
        row.push(val);
      }
    }
    outputRows.push(row);
  }
  
  // スプレッドシートに一括書き込み
  var targetRange = sheet.getRange(1, 1, outputRows.length, outputRows[0].length);

  // ★ isoDate列・recordedAt列は必ずプレーンテキストとして書き込む（Google Sheetsが
  //   ISO形式の文字列を自動的に日付型セルへ変換してしまい、タイムゾーンの
  //   解釈違いにより日付が前後にズレる（＝月をまたいでしまう）事故を防ぐため）。
  //   setValues() より前に書式を「テキスト」に固定しておく必要がある。
  if (sheetName === "history") {
    ["isoDate", "purchaseDate", "recordedAt"].forEach(function(colName) {
      var colIndex = headers.indexOf(colName) + 1; // 1始まり
      if (colIndex > 0 && outputRows.length > 1) {
        sheet.getRange(2, colIndex, outputRows.length - 1, 1).setNumberFormat("@");
      }
    });
  }

  // ★ lockedMonths の month列（"YYYY-MM"）も同様に、Google Sheetsが日付として
  //   自動解釈・変換してしまうのを防ぐため、書き込み前にテキスト形式へ固定する。
  //   これが無いと "2026-05" が日付型セルに変換され、再読込時に
  //   "2026-05-01 00:00" のような別形式になってしまい、フロントエンドの
  //   完全一致判定（lockedMonths.includes(月)）が常に不一致となり、
  //   「仮締めが保存されない（再起動すると解除されている）」不具合の原因になる。
  if (sheetName === "lockedMonths") {
    var monthColIndex = headers.indexOf("month") + 1;
    if (monthColIndex > 0 && outputRows.length > 1) {
      sheet.getRange(2, monthColIndex, outputRows.length - 1, 1).setNumberFormat("@");
    }
  }

  // ★ products の createdMonth・inactiveMonth・reactivateMonth（いずれも "YYYY-MM"）も
  //   同じ理由でテキスト形式に固定する。固定しないと "2026-07" が日付型セルに変換され、
  //   フロントエンドの完全一致判定（toMonthKey での比較）がズレて、商品の表示開始月・
  //   取り扱い終了月・再開月の判定がおかしくなる不具合につながる。
  if (sheetName === "products") {
    ["createdMonth", "inactiveMonth", "reactivateMonth"].forEach(function(colName) {
      var colIndex = headers.indexOf(colName) + 1;
      if (colIndex > 0 && outputRows.length > 1) {
        sheet.getRange(2, colIndex, outputRows.length - 1, 1).setNumberFormat("@");
      }
    });
  }

    targetRange.setValues(outputRows);
  } finally {
    lock.releaseLock();
  }
}

/**
 * 画面から送られた履歴データを「history」シートに追加または上書きする関数
 * ★ 修正: 同一IDの行が既に存在する場合は上書き、存在しない場合のみ追記
 *   ヘッダー順: id | date | isoDate | purchaseDate | product | productCategory | type | qty | price | recordedAt
 */
function addHistoryRow(cleanRow) {
  validateHistoryRow_(cleanRow);
  var lock = acquireDocumentLock_();
  try {
    var ss = SpreadsheetApp.getActiveSpreadsheet();
    var sheet = ss.getSheetByName('history');
    if (!sheet) {
      throw new Error("「history」シートが見つかりません。");
    }
    
    // シートのヘッダー順に合わせて配列に並び替える
    // [ID, 日時, ISO日時, 仕入れた日付, 商品名, 区分, 種別, 数量, 単価, 実処理日時]
    var rowData = [
      cleanRow.id,
      cleanRow.date,
      cleanRow.isoDate,
      cleanRow.purchaseDate || "",
      cleanRow.product,
      cleanRow.productCategory,
      cleanRow.type,
      cleanRow.qty,
      cleanRow.price,
      cleanRow.recordedAt
    ];
    
    // ★ 同一IDの行を検索して上書き。なければ末尾に追記。
    var dataRange = sheet.getDataRange();
    var values = dataRange.getValues();

    // isoDate列（3列目）・purchaseDate列（4列目）・recordedAt列（10列目）は必ずプレーンテキストとして書き込む。
    // （Google Sheetsが文字列を自動的に日付型セルへ変換し、タイムゾーンの
    //   解釈違いで日付が前後にズレて月をまたいでしまう事故を防ぐため。
    //   setValues() より前に書式を固定しておく必要がある）
    var ISO_DATE_COL = 3;
    var PURCHASE_DATE_COL = 4;
    var RECORDED_AT_COL = 10;

    // ヘッダー行(index=0)を除いてID列(index=0)を検索
    for (var i = 1; i < values.length; i++) {
      if (String(values[i][0]) === String(cleanRow.id)) {
        // 既存行を上書き（シート行番号は1始まりなので i+1）
        sheet.getRange(i + 1, ISO_DATE_COL, 1, 1).setNumberFormat("@");
        sheet.getRange(i + 1, PURCHASE_DATE_COL, 1, 1).setNumberFormat("@");
        sheet.getRange(i + 1, RECORDED_AT_COL, 1, 1).setNumberFormat("@");
        sheet.getRange(i + 1, 1, 1, rowData.length).setValues([rowData]);
        return true;
      }
    }
    
    // 同一IDが見つからなければ末尾に追記
    var newRowIndex = sheet.getLastRow() + 1;
    sheet.getRange(newRowIndex, ISO_DATE_COL, 1, 1).setNumberFormat("@");
    sheet.getRange(newRowIndex, PURCHASE_DATE_COL, 1, 1).setNumberFormat("@");
    sheet.getRange(newRowIndex, RECORDED_AT_COL, 1, 1).setNumberFormat("@");
    sheet.getRange(newRowIndex, 1, 1, rowData.length).setValues([rowData]);
    return true;
  } catch (e) {
    throw new Error("GAS側での書き込みに失敗しました: " + e.toString());
  } finally {
    lock.releaseLock();
  }
}
