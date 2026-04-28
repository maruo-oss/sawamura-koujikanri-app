/**
 * 工事管理システム - Google Apps Script メインモジュール
 *
 * Repository.gs と連携してSpreadsheetのデータを操作
 */

// ========================================
// 設定
// ========================================

/** スプレッドシートID */
var SPREADSHEET_ID = '1HrUUs54Np0Qq5Ty8Ubb5FMsTtzbVJujYOMFg0k72fwk';

// ========================================
// 出力設定
// ========================================
var OUTPUT_CONFIG = {
  // テンプレートファイルID（Google Drive）
  // 契約協議書テンプレート
  TEMPLATE_CONTRACT_SOLUTION: '1UGFy7KpgfSvwEGxGwL4rA4SLh2CLWA4CNAMP5asZd18',
  // 契約協議書テンプレート（小口用）
  TEMPLATE_CONTRACT_SMALL: '1DtFYchjR1uIu59yiheup8QhVs9JQXzaGiKMWZWkjMPs',
  // 注文協議書テンプレート
  TEMPLATE_ORDER_DOCUMENT: '130Sk3zJ99Vpsxrlrvi1dq3mTXK7NIK_P-kBTOfYYTEw',
  // 注文書テンプレート（Googleドキュメント形式・プレースホルダー置換方式）
  TEMPLATE_ORDER_SHEET: '1PRNVViQAnyP5pW7f-lu-14z8kDzhayGYnC1YynkA1fE',

  // 出力フォルダID（Google Drive）
  OUTPUT_FOLDER_ID: '1lfLfptzcjGF-fhywtjZO-vsvNZ598l7D'
};

/**
 * スプレッドシートIDを取得
 */
function getSpreadsheetId() {
  return SPREADSHEET_ID;
}

// ========================================
// 診断関数（デバッグ用）
// ========================================

/**
 * スプレッドシートの全シートの状態を診断する
 * GASエディタから手動で実行 → ログ（Ctrl+Enter）で結果確認
 */
function diagnoseSpreadsheet() {
  var ss = SpreadsheetApp.openById(SPREADSHEET_ID);
  var sheets = ss.getSheets();

  Logger.log('========== スプレッドシート診断 ==========');
  Logger.log('スプレッドシート名: ' + ss.getName());
  Logger.log('シート数: ' + sheets.length);
  Logger.log('');

  sheets.forEach(function(sheet) {
    var name = sheet.getName();
    var lastRow = sheet.getLastRow();
    var lastCol = sheet.getLastColumn();
    var headers = lastCol > 0 ? sheet.getRange(1, 1, 1, lastCol).getValues()[0] : [];

    Logger.log('--- シート: ' + name + ' ---');
    Logger.log('  データ行数: ' + (lastRow > 0 ? lastRow - 1 : 0) + '行（ヘッダー除く）');
    Logger.log('  カラム数: ' + lastCol);
    Logger.log('  ヘッダー: ' + headers.join(', '));

    // 2行目のサンプルデータ（あれば）
    if (lastRow >= 2 && lastCol > 0) {
      var sample = sheet.getRange(2, 1, 1, Math.min(lastCol, 8)).getValues()[0];
      var sampleStr = headers.slice(0, 8).map(function(h, i) {
        var val = sample[i];
        if (val instanceof Date) {
          return h + '=' + Utilities.formatDate(val, Session.getScriptTimeZone(), 'yyyy-MM-dd') + '(Date型)';
        }
        return h + '=' + String(val).substring(0, 30);
      }).join(' | ');
      Logger.log('  サンプル(1行目): ' + sampleStr);
    }
    Logger.log('');
  });

  // 特に重要なシートのチェック
  Logger.log('========== 重要シートチェック ==========');
  var requiredSheets = ['案件管理', '工事管理', '契約協議書', '発注業者管理', '請求管理', '支払管理', '実行予算', '顧客マスタ', '社員マスタ', '協力会社マスタ', '監査ログ', '添付ファイル管理', '工事台帳', 'マイルストーン管理'];

  requiredSheets.forEach(function(name) {
    var sheet = ss.getSheetByName(name);
    if (sheet) {
      Logger.log('  ✅ ' + name + ' (データ: ' + (sheet.getLastRow() - 1) + '行)');
    } else {
      Logger.log('  ❌ ' + name + ' ← 存在しません！');
    }
  });

  return 'ログを確認してください';
}

/**
 * API関数の動作テスト
 * GASエディタから手動で実行 → ログで結果確認
 */
function diagnoseApi() {
  Logger.log('========== API動作テスト ==========');

  // 1. getConstructions テスト
  Logger.log('--- getConstructions() テスト ---');
  try {
    var constructions = getConstructions();
    Logger.log('  件数: ' + (constructions ? constructions.length : 'null'));
    if (constructions && constructions.length > 0) {
      Logger.log('  1件目: ' + JSON.stringify(constructions[0]).substring(0, 300));
      Logger.log('  工事番号一覧: ' + constructions.map(function(c) { return c['工事番号'] || c.constructionId; }).join(', '));
    } else {
      Logger.log('  ⚠️ データが0件です！');
      // 直接シートから取得してみる
      var raw = ConstructionRepository.findAll();
      Logger.log('  Repository直接: ' + raw.length + '件');
      if (raw.length > 0) {
        Logger.log('  生データ1件目キー: ' + Object.keys(raw[0]).join(', '));
        Logger.log('  生データ1件目: ' + JSON.stringify(raw[0]).substring(0, 300));
      }
    }
  } catch (e) {
    Logger.log('  ❌ エラー: ' + e.message);
    Logger.log('  スタック: ' + e.stack);
  }

  // 2. getGanttChartData テスト
  Logger.log('');
  Logger.log('--- getGanttChartData() テスト ---');
  try {
    var gantt = getGanttChartData({});
    Logger.log('  success: ' + gantt.success);
    Logger.log('  items件数: ' + (gantt.items ? gantt.items.length : 'null'));
    if (gantt.items && gantt.items.length > 0) {
      Logger.log('  1件目: ' + JSON.stringify(gantt.items[0]).substring(0, 300));
    }
    if (gantt.error) Logger.log('  エラー: ' + gantt.error);
  } catch (e) {
    Logger.log('  ❌ エラー: ' + e.message);
    Logger.log('  スタック: ' + e.stack);
  }

  // 3. getMilestones テスト（最初の工事）
  Logger.log('');
  Logger.log('--- getMilestones() テスト ---');
  try {
    var allMs = MilestoneRepository.findAll();
    Logger.log('  マイルストーン全件: ' + allMs.length);
    if (allMs.length > 0) {
      var firstMs = allMs[0];
      Logger.log('  1件目: ' + JSON.stringify(firstMs).substring(0, 200));
      var planDate = firstMs['予定日'];
      Logger.log('  予定日の型: ' + (planDate instanceof Date ? 'Date' : typeof planDate) + ' 値: ' + planDate);
    }
  } catch (e) {
    Logger.log('  ❌ エラー: ' + e.message);
  }

  // 4. 工事台帳ダミーデータの工事番号チェック
  Logger.log('');
  Logger.log('--- 工事台帳⇔工事管理 ID整合性チェック ---');
  try {
    var ledgers = LedgerRepository.findAll();
    Logger.log('  工事台帳: ' + ledgers.length + '件');
    ledgers.forEach(function(l) {
      var cId = l['工事番号'];
      var found = ConstructionRepository.findById(cId);
      Logger.log('  台帳[' + l['台帳ID'] + '] → 工事番号[' + cId + '] → ' + (found ? '✅ 工事あり' : '❌ 工事なし！'));
    });

    var milestones = MilestoneRepository.findAll();
    var msConstructionIds = [];
    milestones.forEach(function(ms) {
      var cId = ms['工事番号'];
      if (msConstructionIds.indexOf(cId) === -1) msConstructionIds.push(cId);
    });
    Logger.log('');
    Logger.log('  マイルストーンが参照する工事番号:');
    msConstructionIds.forEach(function(cId) {
      var found = ConstructionRepository.findById(cId);
      Logger.log('  MS→工事番号[' + cId + '] → ' + (found ? '✅ 工事あり' : '❌ 工事なし！'));
    });
  } catch (e) {
    Logger.log('  ❌ エラー: ' + e.message);
  }

  return 'ログを確認してください';
}

// ========================================
// 初期化・セットアップ関数
// ========================================

/**
 * 新機能用のシートを作成し、ダミーデータを投入する
 * GASエディタから手動で実行してください
 */
function setupNewFeatureSheets() {
  var ss = SpreadsheetApp.openById(SPREADSHEET_ID);
  var results = [];

  // 1. 監査ログシート
  results.push(createAuditLogSheet(ss));

  // 2. 添付ファイル管理シート
  results.push(createAttachmentSheet(ss));

  // 3. 工事台帳シート
  results.push(createLedgerSheet(ss));

  // 4. マイルストーン管理シート
  results.push(createMilestoneSheet(ss));

  // 5. 既存シートに削除フラグカラムを追加
  results.push(addSoftDeleteColumns(ss));

  // 6. 工事管理シートに進捗カラムを追加
  results.push(addProgressColumns(ss));

  Logger.log('=== セットアップ完了 ===');
  results.forEach(function(r) {
    Logger.log(r);
  });

  return results;
}

/**
 * 監査ログシートを作成
 */
function createAuditLogSheet(ss) {
  var sheetName = '監査ログ';
  var existing = ss.getSheetByName(sheetName);
  if (existing) {
    return sheetName + ': 既に存在します（スキップ）';
  }

  var sheet = ss.insertSheet(sheetName);

  // ヘッダー
  var headers = [
    'ログID', 'タイムスタンプ', 'ユーザーEmail', 'ユーザー名', '部門',
    '操作種別', '対象シート', '対象ID', '変更前JSON', '変更後JSON', '変更フィールド'
  ];
  sheet.getRange(1, 1, 1, headers.length).setValues([headers]);
  sheet.getRange(1, 1, 1, headers.length).setFontWeight('bold').setBackground('#4a86e8').setFontColor('white');

  // ダミーデータ
  var now = new Date();
  var dummyData = [];
  var operations = ['CREATE', 'UPDATE', 'DELETE', 'EXPORT'];
  var sheets = ['案件管理', '契約協議書', '工事管理', '発注業者管理'];

  for (var i = 0; i < 10; i++) {
    var opDate = new Date(now.getTime() - i * 3600000); // 1時間ずつ前
    dummyData.push([
      'LOG-' + Utilities.formatDate(opDate, Session.getScriptTimeZone(), 'yyyyMMdd-HHmmss') + '-' + String(i).padStart(3, '0'),
      Utilities.formatDate(opDate, Session.getScriptTimeZone(), 'yyyy-MM-dd HH:mm:ss'),
      'user' + (i % 3 + 1) + '@example.com',
      ['山田太郎', '佐藤花子', '田中一郎'][i % 3],
      'ソリューション事業部',
      operations[i % 4],
      sheets[i % 4],
      'PJ-2024-000' + (i + 1),
      i % 4 === 1 ? '{"status":"進行中"}' : '',
      i % 4 === 1 ? '{"status":"完了"}' : '{"案件名":"テスト案件' + i + '"}',
      i % 4 === 1 ? 'status' : '案件名'
    ]);
  }

  if (dummyData.length > 0) {
    sheet.getRange(2, 1, dummyData.length, headers.length).setValues(dummyData);
  }

  // 列幅調整
  sheet.setColumnWidth(1, 200);  // ログID
  sheet.setColumnWidth(2, 150);  // タイムスタンプ
  sheet.setColumnWidth(9, 300);  // 変更前JSON
  sheet.setColumnWidth(10, 300); // 変更後JSON

  return sheetName + ': 作成完了（' + dummyData.length + '件のダミーデータ）';
}

/**
 * 添付ファイル管理シートを作成
 */
function createAttachmentSheet(ss) {
  var sheetName = '添付ファイル管理';
  var existing = ss.getSheetByName(sheetName);
  if (existing) {
    return sheetName + ': 既に存在します（スキップ）';
  }

  var sheet = ss.insertSheet(sheetName);

  // ヘッダー
  var headers = [
    'ファイルID', 'GoogleDriveファイルID', 'ファイル名', 'MIMEタイプ', 'ファイルサイズ',
    '関連エンティティ種別', '関連エンティティID', 'カテゴリ', 'アップロード日時', 'アップロード者',
    '削除フラグ', '削除日時', '削除者'
  ];
  sheet.getRange(1, 1, 1, headers.length).setValues([headers]);
  sheet.getRange(1, 1, 1, headers.length).setFontWeight('bold').setBackground('#4a86e8').setFontColor('white');

  // ダミーデータ
  var now = new Date();
  var categories = ['契約書', '図面', '写真', '見積書', '報告書', 'その他'];
  var entityTypes = ['project', 'construction', 'contract', 'order'];
  var mimeTypes = ['application/pdf', 'image/jpeg', 'image/png', 'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet'];
  var fileNames = ['契約書_2024.pdf', '現場写真_001.jpg', '設計図面.png', '見積書_v1.xlsx', '報告書_最終.pdf'];

  var dummyData = [];
  for (var i = 0; i < 8; i++) {
    var uploadDate = new Date(now.getTime() - i * 86400000); // 1日ずつ前
    dummyData.push([
      'ATT-' + Utilities.formatDate(uploadDate, Session.getScriptTimeZone(), 'yyyyMMdd') + '-' + String(i + 1).padStart(6, '0'),
      '1abc' + String(i).padStart(20, 'x'), // ダミーのDrive ID
      fileNames[i % 5],
      mimeTypes[i % 4],
      Math.floor(Math.random() * 5000000) + 100000, // 100KB〜5MB
      entityTypes[i % 4],
      i % 4 === 0 ? 'PJ-2024-000' + (i + 1) : (i % 4 === 1 ? 'C-0000' + (i + 1) : 'CT-2024-000' + (i + 1)),
      categories[i % 6],
      Utilities.formatDate(uploadDate, Session.getScriptTimeZone(), 'yyyy-MM-dd HH:mm:ss'),
      ['山田太郎', '佐藤花子', '田中一郎'][i % 3],
      '', '', '' // 削除フラグ関連は空
    ]);
  }

  if (dummyData.length > 0) {
    sheet.getRange(2, 1, dummyData.length, headers.length).setValues(dummyData);
  }

  sheet.setColumnWidth(1, 180);
  sheet.setColumnWidth(2, 200);
  sheet.setColumnWidth(3, 200);

  return sheetName + ': 作成完了（' + dummyData.length + '件のダミーデータ）';
}

/**
 * 工事台帳シートを作成
 */
function createLedgerSheet(ss) {
  var sheetName = '工事台帳';
  var existing = ss.getSheetByName(sheetName);
  if (existing) {
    return sheetName + ': 既に存在します（スキップ）';
  }

  var sheet = ss.insertSheet(sheetName);

  // ヘッダー
  var headers = [
    '台帳ID', '工事番号', '案件ID', '作成日', 'ステータス', '更新日時', '備考',
    '削除フラグ', '削除日時', '削除者'
  ];
  sheet.getRange(1, 1, 1, headers.length).setValues([headers]);
  sheet.getRange(1, 1, 1, headers.length).setFontWeight('bold').setBackground('#4a86e8').setFontColor('white');

  // ダミーデータ（工事管理シートの工事に対応）
  var now = new Date();
  var statuses = ['作成中', '確認中', '確定'];
  var dummyData = [];

  for (var i = 0; i < 5; i++) {
    var createDate = new Date(now.getTime() - i * 7 * 86400000); // 1週間ずつ前
    dummyData.push([
      'LED-2024-' + String(i + 1).padStart(4, '0'),
      'C-0000' + (i + 1),
      'PJ-2024-000' + (i + 1),
      Utilities.formatDate(createDate, Session.getScriptTimeZone(), 'yyyy-MM-dd'),
      statuses[i % 3],
      Utilities.formatDate(now, Session.getScriptTimeZone(), 'yyyy-MM-dd HH:mm:ss'),
      i === 0 ? '初回作成' : '',
      '', '', ''
    ]);
  }

  if (dummyData.length > 0) {
    sheet.getRange(2, 1, dummyData.length, headers.length).setValues(dummyData);
  }

  sheet.setColumnWidth(1, 150);
  sheet.setColumnWidth(2, 100);
  sheet.setColumnWidth(3, 120);

  return sheetName + ': 作成完了（' + dummyData.length + '件のダミーデータ）';
}

/**
 * マイルストーン管理シートを作成
 */
function createMilestoneSheet(ss) {
  var sheetName = 'マイルストーン管理';
  var existing = ss.getSheetByName(sheetName);
  if (existing) {
    return sheetName + ': 既に存在します（スキップ）';
  }

  var sheet = ss.insertSheet(sheetName);

  // ヘッダー
  var headers = [
    'マイルストーンID', '工事番号', '名称', '予定日', '実績日', 'ステータス'
  ];
  sheet.getRange(1, 1, 1, headers.length).setValues([headers]);
  sheet.getRange(1, 1, 1, headers.length).setFontWeight('bold').setBackground('#4a86e8').setFontColor('white');

  // ダミーデータ
  var now = new Date();
  var milestoneNames = ['着工', '基礎工事完了', '中間検査', '竣工検査', '引渡し'];
  var dummyData = [];
  var msId = 1;

  // 工事ごとにマイルストーンを作成
  for (var c = 1; c <= 3; c++) {
    var constructionId = 'C-0000' + c;
    var baseDate = new Date(now.getTime() - (3 - c) * 30 * 86400000); // 工事ごとに1ヶ月ずらす

    for (var m = 0; m < milestoneNames.length; m++) {
      var planDate = new Date(baseDate.getTime() + m * 14 * 86400000); // 2週間ずつ
      var isCompleted = c === 1 || (c === 2 && m < 3);
      var isDelayed = !isCompleted && planDate < now;

      dummyData.push([
        'MS-' + String(msId++).padStart(6, '0'),
        constructionId,
        milestoneNames[m],
        Utilities.formatDate(planDate, Session.getScriptTimeZone(), 'yyyy-MM-dd'),
        isCompleted ? Utilities.formatDate(new Date(planDate.getTime() + Math.random() * 3 * 86400000), Session.getScriptTimeZone(), 'yyyy-MM-dd') : '',
        isCompleted ? '完了' : (isDelayed ? '遅延' : '予定')
      ]);
    }
  }

  if (dummyData.length > 0) {
    sheet.getRange(2, 1, dummyData.length, headers.length).setValues(dummyData);
  }

  sheet.setColumnWidth(1, 150);
  sheet.setColumnWidth(2, 100);
  sheet.setColumnWidth(3, 120);
  sheet.setColumnWidth(4, 100);
  sheet.setColumnWidth(5, 100);

  return sheetName + ': 作成完了（' + dummyData.length + '件のダミーデータ）';
}

/**
 * 既存シートに削除フラグカラムを追加
 */
function addSoftDeleteColumns(ss) {
  var targetSheets = ['案件管理', '契約協議書', '工事管理', '発注業者管理', '請求管理', '支払管理', '実行予算'];
  var columnsToAdd = ['削除フラグ', '削除日時', '削除者'];
  var results = [];

  targetSheets.forEach(function(sheetName) {
    var sheet = ss.getSheetByName(sheetName);
    if (!sheet) {
      results.push(sheetName + ': シートが見つかりません');
      return;
    }

    var headers = sheet.getRange(1, 1, 1, sheet.getLastColumn()).getValues()[0];
    var addedColumns = [];

    columnsToAdd.forEach(function(colName) {
      if (headers.indexOf(colName) === -1) {
        var newCol = sheet.getLastColumn() + 1;
        sheet.getRange(1, newCol).setValue(colName);
        addedColumns.push(colName);
      }
    });

    if (addedColumns.length > 0) {
      results.push(sheetName + ': カラム追加 [' + addedColumns.join(', ') + ']');
    } else {
      results.push(sheetName + ': 削除フラグカラム既存（スキップ）');
    }
  });

  return results.join('\n');
}

/**
 * 工事管理シートに進捗カラムを追加
 */
function addProgressColumns(ss) {
  var sheetName = '工事管理';
  var sheet = ss.getSheetByName(sheetName);
  if (!sheet) {
    return sheetName + ': シートが見つかりません';
  }

  var headers = sheet.getRange(1, 1, 1, sheet.getLastColumn()).getValues()[0];
  var columnsToAdd = ['進捗率', '進捗更新日', '進捗メモ'];
  var addedColumns = [];

  columnsToAdd.forEach(function(colName) {
    if (headers.indexOf(colName) === -1) {
      var newCol = sheet.getLastColumn() + 1;
      sheet.getRange(1, newCol).setValue(colName);
      addedColumns.push(colName);
    }
  });

  // ダミーの進捗データを入力（既存データがある行に）
  if (addedColumns.indexOf('進捗率') !== -1) {
    var lastRow = sheet.getLastRow();
    if (lastRow > 1) {
      var progressCol = sheet.getRange(1, 1, 1, sheet.getLastColumn()).getValues()[0].indexOf('進捗率') + 1;
      for (var row = 2; row <= lastRow; row++) {
        sheet.getRange(row, progressCol).setValue(Math.floor(Math.random() * 100));
      }
    }
  }

  if (addedColumns.length > 0) {
    return sheetName + ': カラム追加 [' + addedColumns.join(', ') + '] + ダミー進捗率設定';
  } else {
    return sheetName + ': 進捗カラム既存（スキップ）';
  }
}

/**
 * 全ての新機能シートを削除（やり直し用）
 * 注意: データが消えるので本番環境では使用しないこと
 */
function deleteNewFeatureSheets() {
  var ss = SpreadsheetApp.openById(SPREADSHEET_ID);
  var sheetsToDelete = ['監査ログ', '添付ファイル管理', '工事台帳', 'マイルストーン管理'];
  var results = [];

  sheetsToDelete.forEach(function(sheetName) {
    var sheet = ss.getSheetByName(sheetName);
    if (sheet) {
      ss.deleteSheet(sheet);
      results.push(sheetName + ': 削除しました');
    } else {
      results.push(sheetName + ': 存在しません（スキップ）');
    }
  });

  Logger.log('=== シート削除完了 ===');
  results.forEach(function(r) { Logger.log(r); });
  return results;
}

/**
 * 支払管理シートに新カラム（決済条件・安推協対象外）を追加
 * GASエディタから一度だけ手動実行してください
 */
function migratePaymentColumns() {
  var ss = SpreadsheetApp.openById(SPREADSHEET_ID);
  var results = [];

  // 支払管理シートに新カラムを追加
  var paymentSheet = ss.getSheetByName('支払管理');
  if (paymentSheet) {
    var lastCol = paymentSheet.getLastColumn();
    var headers = paymentSheet.getRange(1, 1, 1, lastCol).getValues()[0];
    var newColumns = [
      { name: '決済条件', defaultValue: '通常' },
      { name: '安推協対象外', defaultValue: 'いいえ' }
    ];

    newColumns.forEach(function(col) {
      if (headers.indexOf(col.name) === -1) {
        var newCol = paymentSheet.getLastColumn() + 1;
        paymentSheet.getRange(1, newCol).setValue(col.name);
        paymentSheet.getRange(1, newCol).setFontWeight('bold').setBackground('#4a86e8').setFontColor('white');
        // 既存データ行にデフォルト値を設定
        var lastRow = paymentSheet.getLastRow();
        if (lastRow > 1) {
          var range = paymentSheet.getRange(2, newCol, lastRow - 1, 1);
          var defaults = [];
          for (var i = 0; i < lastRow - 1; i++) {
            defaults.push([col.defaultValue]);
          }
          range.setValues(defaults);
        }
        results.push('支払管理: 「' + col.name + '」カラム追加（デフォルト: ' + col.defaultValue + '）');
      } else {
        results.push('支払管理: 「' + col.name + '」は既に存在（スキップ）');
      }
    });
  }

  SpreadsheetApp.flush();
  Logger.log('=== 支払カラムマイグレーション完了 ===');
  results.forEach(function(r) { Logger.log(r); });
  return results;
}

/**
 * スプレッドシートオブジェクトのキャッシュ
 * 同一リクエスト内で何度もopenById()を呼ぶのを防ぐ
 */
var _ssCache = null;

function getSpreadsheet() {
  if (!_ssCache) {
    _ssCache = SpreadsheetApp.openById(SPREADSHEET_ID);
  }
  return _ssCache;
}

// ========================================
// Webアプリ エントリポイント
// ========================================

/**
 * HTTP GETリクエストのハンドラ
 */
function doGet(e) {
  var template = HtmlService.createTemplateFromFile('index');

  return template.evaluate()
    .setTitle('工事管理システム')
    .setXFrameOptionsMode(HtmlService.XFrameOptionsMode.ALLOWALL)
    .addMetaTag('viewport', 'width=device-width, initial-scale=1');
}

/**
 * HTMLテンプレートでファイルをインクルードするためのヘルパー関数
 */
function include(filename) {
  return HtmlService.createHtmlOutputFromFile(filename).getContent();
}

// ========================================
// ユーザー・認証 API
// ========================================

/**
 * 現在のログインユーザー情報を取得
 * 社員マスタと連携して詳細情報を返す
 */
function getCurrentUser() {
  var email = Session.getActiveUser().getEmail();

  // 社員マスタから検索
  try {
    var employee = EmployeeRepository.findByEmail(email);
    if (employee) {
      return {
        id: employee['社員番号'],
        name: employee['氏名'],
        email: email,
        department: employee['部門'],
        group: employee['所属グループ(主務)'],
        role: employee['承認権限レベル'] || 'staff',
        managerEmail: employee['上長メールアドレス']
      };
    }
  } catch (e) {
    Logger.log('社員マスタ検索エラー: ' + e.message);
  }

  // 見つからない場合はデフォルト
  var name = email.split('@')[0];
  return {
    id: name,
    name: name,
    email: email,
    department: 'ソリューション事業部',
    role: 'staff'
  };
}

// ========================================
// 権限管理
// ========================================

/**
 * 権限レベル定義
 * 数値が大きいほど権限が高い
 */
var PERMISSION_LEVELS = {
  readonly: 1,  // 閲覧専用
  staff: 2,     // 一般社員（自部門データのCRUD）
  manager: 3,   // 管理職（承認権限、部門データ閲覧）
  admin: 4      // システム管理者（全機能アクセス可）
};

/**
 * 機能別の必要権限定義
 */
var PERMISSIONS = {
  // 工事管理
  'project.view': 'readonly',
  'project.create': 'staff',
  'project.update': 'staff',
  'project.delete': 'manager',

  // 契約協議
  'contract.view': 'readonly',
  'contract.create': 'staff',
  'contract.update': 'staff',
  'contract.approve': 'manager',
  'contract.delete': 'manager',

  // 発注管理
  'order.view': 'readonly',
  'order.create': 'staff',
  'order.update': 'staff',
  'order.approve': 'manager',
  'order.delete': 'manager',

  // 工事管理
  'construction.view': 'readonly',
  'construction.create': 'staff',
  'construction.update': 'staff',
  'construction.delete': 'manager',

  // 請求管理
  'invoice.view': 'readonly',
  'invoice.create': 'staff',
  'invoice.update': 'staff',
  'invoice.delete': 'manager',

  // 支払管理
  'payment.view': 'readonly',
  'payment.create': 'staff',
  'payment.update': 'staff',
  'payment.approve': 'manager',
  'payment.delete': 'manager',

  // 実行予算
  'budget.view': 'readonly',
  'budget.create': 'staff',
  'budget.update': 'staff',
  'budget.delete': 'manager',

  // マスタ管理
  'master.view': 'staff',
  'master.create': 'admin',
  'master.update': 'manager',
  'master.delete': 'admin',

  // 監査ログ
  'audit.view': 'manager',

  // アーカイブ
  'archive.view': 'manager',
  'archive.restore': 'manager',
  'archive.permanentDelete': 'admin',

  // インポート
  'import.execute': 'manager',

  // 工事台帳
  'ledger.view': 'readonly',
  'ledger.create': 'staff',
  'ledger.export': 'staff',

  // スケジュール
  'schedule.view': 'readonly',
  'schedule.update': 'staff'
};

/**
 * 権限チェック
 * @param {string} permission - 権限名（例: 'project.create'）
 * @param {Object} user - ユーザー情報（省略時は現在のユーザー）
 * @returns {boolean} 権限があればtrue
 */
function checkPermission(permission, user) {
  if (!user) {
    user = getCurrentUser();
  }

  var userRole = user.role || 'staff';
  var requiredRole = PERMISSIONS[permission];

  if (!requiredRole) {
    Logger.log('権限定義が見つかりません: ' + permission);
    return false;
  }

  var userLevel = PERMISSION_LEVELS[userRole] || 0;
  var requiredLevel = PERMISSION_LEVELS[requiredRole] || 0;

  return userLevel >= requiredLevel;
}

/**
 * 権限チェック付き関数実行ラッパー
 * @param {string} permission - 必要な権限
 * @param {Function} fn - 実行する関数
 * @returns {any} 関数の戻り値または権限エラー
 */
function executeWithPermission(permission, fn) {
  if (!checkPermission(permission)) {
    return {
      success: false,
      error: 'この操作を実行する権限がありません',
      permissionDenied: true
    };
  }

  try {
    return fn();
  } catch (e) {
    Logger.log('権限付き実行エラー: ' + e.message);
    return {
      success: false,
      error: e.message
    };
  }
}

/**
 * ユーザーの権限一覧を取得
 * フロントエンドでのUI制御に使用
 * @returns {Object} 権限マップ
 */
function getUserPermissions() {
  var user = getCurrentUser();
  var permissions = {};

  Object.keys(PERMISSIONS).forEach(function(key) {
    permissions[key] = checkPermission(key, user);
  });

  return {
    role: user.role,
    permissions: permissions
  };
}

/**
 * アプリ初期化用一括データ取得
 * google.script.runの往復回数を1回に削減して高速化
 */
function initializeApp() {
  try {
    var user = getCurrentUser();

    // 権限を計算
    var permissions = {};
    Object.keys(PERMISSIONS).forEach(function(key) {
      permissions[key] = checkPermission(key, user);
    });

    // 工事一覧（getProjects相当）
    var projects = [];
    try {
      projects = getProjects({ department: user.department });
    } catch (e) {
      Logger.log('initializeApp - 工事取得エラー: ' + e.message);
    }

    // 工事一覧
    var constructions = [];
    try {
      constructions = getConstructions();
    } catch (e) {
      Logger.log('initializeApp - 工事一覧取得エラー: ' + e.message);
    }

    // 契約協議一覧
    var contracts = [];
    try {
      contracts = getContracts('ソリューション事業部');
    } catch (e) {
      Logger.log('initializeApp - 契約取得エラー: ' + e.message);
    }

    // マスタデータ
    var customers = [];
    var vendors = [];
    var employees = [];
    try { customers = getCustomersLight(); } catch (e) { Logger.log('initializeApp - 顧客取得エラー: ' + e.message); }
    try { vendors = getVendorsLight(); } catch (e) { Logger.log('initializeApp - 協力会社取得エラー: ' + e.message); }
    try { employees = getEmployees(); } catch (e) { Logger.log('initializeApp - 社員取得エラー: ' + e.message); }

    return {
      currentUser: user,
      permissions: { role: user.role, permissions: permissions },
      projects: projects,
      constructions: constructions,
      contracts: contracts,
      customers: customers,
      vendors: vendors,
      employees: employees
    };
  } catch (e) {
    Logger.log('initializeApp エラー: ' + e.message);
    return { error: e.message };
  }
}

// ========================================
// 工事管理 API
// ========================================

/**
 * 工事一覧を取得
 * スプレッドシートのカラム名をそのまま返す（日本語カラム名）
 * 顧客IDは顧客名_表示フィールドに名前変換して追加
 */
function getProjects(filter) {
  try {
    // 論理削除済みを除外して取得（getAllDataExcludeDeletedを使用）
    var allData = ProjectRepository.findAll();

    // マスタルックアップを取得（キャッシュ利用）
    var customerMap = MasterLookup.getCustomerMap();

    var projects = allData.map(function(obj) {
      // Date型は文字列に変換（シリアライズ問題対策）
      for (var key in obj) {
        if (obj.hasOwnProperty(key) && obj[key] instanceof Date) {
          obj[key] = Utilities.formatDate(obj[key], Session.getScriptTimeZone(), 'yyyy-MM-dd');
        }
      }

      // ID→名前変換を付与（「顧客名」列には実際はIDが格納されている）
      var customerId = obj['顧客名'];
      obj['顧客ID'] = customerId;
      obj['顧客名_表示'] = customerMap[customerId] || customerId;

      return obj;
    });

    // 新しいデータを先頭に表示（スプレッドシートは末尾追加のため逆順にする）
    projects.reverse();

    return projects;
  } catch (e) {
    Logger.log('工事取得エラー: ' + e.message);
    return [];
  }
}

/**
 * 工事を1件取得
 */
function getProjectById(projectId) {
  try {
    var row = ProjectRepository.findById(projectId);
    if (!row) return null;

    return {
      projectId: row['案件ID'],
      projectName: row['案件名'],
      customerName: row['顧客名'],
      contractAmount: row['契約金額'] || 0,
      status: row['案件ステータス'],
      managerName: row['担当者名'],
      department: row['部署'],
      category: row['振り分けカテゴリー'],
      startDate: formatDateForApi(row['着工日']),
      endDate: formatDateForApi(row['竣工予定日']),
      contractDate: formatDateForApi(row['契約日'])
    };
  } catch (e) {
    Logger.log('工事取得エラー: ' + e.message);
    return null;
  }
}

/**
 * 工事を保存
 */
function saveProject(data) {
  try {
    // フロントエンドのフィールド名をシートのカラム名に変換
    var sheetData = {
      '案件ID': data.projectId,
      '案件名': data.projectName,
      '顧客名': data.customerName,
      '契約金額': data.contractAmount,
      '案件ステータス': data.status,
      '担当者名': data.managerName,
      '部署': data.department,
      '振り分けカテゴリー': data.category,
      '着工日': data.startDate,
      '竣工予定日': data.endDate,
      '契約日': data.contractDate
    };

    var result = ProjectRepository.save(sheetData);
    return { success: true, id: result['案件ID'] };
  } catch (e) {
    Logger.log('工事保存エラー: ' + e.message);
    return { success: false, error: e.message };
  }
}

/**
 * 工事を論理削除（カスケード削除付き）
 * 関連する発注・見積もり・マイルストーン・添付ファイルも連鎖削除する
 */
function deleteConstruction(constructionId) {
  try {
    if (!constructionId) {
      return { success: false, error: '工事番号が指定されていません' };
    }
    var warnings = [];

    // 1. 関連する発注を論理削除
    var orders = OrderRepository.findByConstructionId(constructionId);
    orders.forEach(function(order) {
      try { OrderRepository.softDelete(order['発注ID']); }
      catch(e) { warnings.push('発注削除失敗: ' + order['発注ID']); }
    });

    // 2. 関連する見積もりを物理削除
    var quotes = QuoteRepository.findByConstructionId(constructionId);
    quotes.forEach(function(q) {
      try { QuoteRepository.delete(q['見積もり詳細ID']); }
      catch(e) { warnings.push('見積削除失敗: ' + q['見積もり詳細ID']); }
    });

    // 3. 関連するマイルストーンを物理削除
    var milestones = MilestoneRepository.findByConstructionId(constructionId);
    milestones.forEach(function(m) {
      try { MilestoneRepository.delete(m['マイルストーンID']); }
      catch(e) { warnings.push('MS削除失敗: ' + m['マイルストーンID']); }
    });

    // 4. 関連する添付ファイルを論理削除
    var attachments = AttachmentRepository.findByEntity('construction', constructionId);
    attachments.forEach(function(att) {
      try { AttachmentRepository.softDelete(att['ファイルID']); }
      catch(e) { warnings.push('添付削除失敗: ' + att['ファイルID']); }
    });

    // 5. 工事本体を論理削除
    ConstructionRepository.softDelete(constructionId);

    Logger.log('工事削除成功(カスケード): ' + constructionId);
    return { success: true, id: constructionId, warnings: warnings.length > 0 ? warnings : undefined };
  } catch (e) {
    Logger.log('工事削除エラー: ' + e.message);
    return { success: false, error: e.message };
  }
}

/**
 * 案件を論理削除（カスケード削除付き）
 * 関連する工事（+その子データ）・請求・予算も連鎖削除する
 */
function deleteProject(projectId) {
  try {
    if (!projectId) {
      return { success: false, error: '案件IDが指定されていません' };
    }
    var warnings = [];

    // 1. 関連する工事をカスケード削除（発注・見積もり等も連鎖）
    var constructions = ConstructionRepository.findByProjectId(projectId);
    constructions.forEach(function(c) {
      var result = deleteConstruction(c['工事番号']);
      if (result.warnings) { warnings = warnings.concat(result.warnings); }
      if (!result.success) { warnings.push('工事削除失敗: ' + c['工事番号']); }
    });

    // 2. 関連する請求を論理削除
    var invoices = InvoiceRepository.findByProjectId(projectId);
    invoices.forEach(function(inv) {
      try { InvoiceRepository.softDelete(inv['請求ID']); }
      catch(e) { warnings.push('請求削除失敗: ' + inv['請求ID']); }
    });

    // 3. 関連する予算を論理削除
    var budgets = BudgetRepository.findByProjectId(projectId);
    budgets.forEach(function(b) {
      try { BudgetRepository.softDelete(b['予算ID']); }
      catch(e) { warnings.push('予算削除失敗: ' + b['予算ID']); }
    });

    // 4. 案件本体を論理削除
    ProjectRepository.softDelete(projectId);

    Logger.log('案件削除成功(カスケード): ' + projectId);
    return { success: true, id: projectId, warnings: warnings.length > 0 ? warnings : undefined };
  } catch (e) {
    Logger.log('案件削除エラー: ' + e.message);
    return { success: false, error: e.message };
  }
}

/**
 * 契約協議書を論理削除（カスケード削除付き）
 * 関連する工事（+その子データ）も連鎖削除する
 */
function deleteContract(contractId) {
  try {
    if (!contractId) {
      return { success: false, error: '協議書IDが指定されていません' };
    }
    var warnings = [];

    // 1. 関連する工事を全件取得してカスケード削除
    var constructions = ConstructionRepository.findAllByContractId(contractId);
    constructions.forEach(function(c) {
      var result = deleteConstruction(c['工事番号']);
      if (result.warnings) { warnings = warnings.concat(result.warnings); }
      if (!result.success) { warnings.push('工事削除失敗: ' + c['工事番号']); }
    });

    // 2. 契約本体を論理削除
    ContractRepository.softDelete(contractId);

    Logger.log('契約協議書削除成功(カスケード): ' + contractId);
    return { success: true, id: contractId, warnings: warnings.length > 0 ? warnings : undefined };
  } catch (e) {
    Logger.log('契約協議書削除エラー: ' + e.message);
    return { success: false, error: e.message };
  }
}

// ========================================
// 工事管理 API
// ========================================

/**
 * 工事一覧を取得（発注フォームのドロップダウン用）
 * 契約協議書から自動登録された工事を含む全工事を返す
 * @param {Object} filter - フィルター条件（オプション）
 * @returns {Array} 工事一覧
 */
function getConstructions(filter) {
  try {
    var constructions = ConstructionRepository.findAll();

    // 新しいデータを先頭に表示
    return constructions.map(function(row) {
      return {
        // 英語キー（既存互換）
        constructionId: row['工事番号'],
        constructionName: row['工事名'],
        contractId: row['関連契約協議書ID'] || '',
        projectId: row['関連案件ID'] || '',
        location: row['工事場所'] || '',
        manager: String(row['工事担当社員'] || ''),
        startDate: formatDateForApi(row['工期始']),
        endDate: formatDateForApi(row['工期終']),
        status: row['進捗ステータス'] || '未着工',
        department: row['部署'] || '',
        locked: row['ロック'] === 'TRUE' || row['ロック'] === true,
        // 日本語キー（新規コンポーネント用）
        '工事番号': row['工事番号'],
        '工事名称': row['工事名'] || row['工事名称'] || '',
        '関連案件ID': row['関連案件ID'] || '',
        '関連契約協議書ID': row['関連契約協議書ID'] || '',
        'ステータス': row['進捗ステータス'] || '未着工',
        '着工日': formatDateForApi(row['工期始']) || formatDateForApi(row['着工日']) || '',
        '完工日': formatDateForApi(row['工期終']) || formatDateForApi(row['完工日']) || '',
        '進捗率': row['進捗率'] || 0
      };
    }).reverse();
  } catch (e) {
    Logger.log('工事一覧取得エラー: ' + e.message);
    return [];
  }
}

/**
 * 工事一覧を取得（工事IDで絞り込み）
 * 工事管理シート12カラムすべてを返す
 */
function getConstructionsByProjectId(projectId) {
  Logger.log('getConstructionsByProjectId開始 - projectId: ' + projectId);
  try {
    var constructions = ConstructionRepository.findByProjectId(projectId);
    Logger.log('工事取得: ' + constructions.length + '件');

    return constructions.map(function(row) {
      return {
        constructionId: row['工事番号'] || '',
        projectId: row['関連案件ID'] || '',
        contractId: row['関連契約協議書ID'] || '',
        recordDate: formatDateForApi(row['記録日']),
        manager: row['工事担当社員'] || '',
        salesPerson: row['営業担当社員'] || '',
        constructionName: row['工事名'] || '',
        location: row['工事場所'] || '',
        workType: row['工種・工事範囲'] || '',
        startDate: formatDateForApi(row['工期始']),
        endDate: formatDateForApi(row['工期終']),
        status: row['進捗ステータス'] || '',
        orderAgreementNo: row['作業所注文協議書ＮＯ．'] || '',
        submissionDate: formatDateForApi(row['提出日']),
        locked: row['ロック'] === 'TRUE' || row['ロック'] === true
      };
    }).reverse();
  } catch (e) {
    Logger.log('工事取得エラー: ' + e.message);
    return [];
  }
}

/**
 * 工事を保存（新規/更新）
 * 工事管理シート12カラムに対応
 */
function saveConstruction(data) {
  Logger.log('saveConstruction開始');
  try {
    // ロックチェック：既存工事の場合、ロック中なら編集拒否
    if (data.constructionId) {
      var existing = ConstructionRepository.findById(data.constructionId);
      if (existing && (existing['ロック'] === 'TRUE' || existing['ロック'] === true)) {
        Logger.log('ロック中の工事への保存を拒否: ' + data.constructionId);
        return { success: false, error: '決算確定済のためロック中です。編集できません。' };
      }
    }
    // フロントエンドのキー名をスプレッドシートのカラム名に変換
    // 工事管理シートの全12カラム: 工事番号, 関連案件ID, 記録日, 工事担当社員, 営業担当社員, 工事名, 工種・工事範囲, 工期始, 工期終, 進捗ステータス, 作業所注文協議書ＮＯ．, 提出日
    var constructionData = {
      '工事番号': data.constructionId || '',
      '関連案件ID': data.projectId || '',
      '記録日': data.recordDate || '',
      '工事担当社員': data.manager || '',
      '営業担当社員': data.salesPerson || '',
      '工事名': data.constructionName || '',
      '工種・工事範囲': data.workType || '',
      '工期始': data.startDate || '',
      '工期終': data.endDate || '',
      '進捗ステータス': data.status || '未着手',
      '作業所注文協議書ＮＯ．': data.orderAgreementNo || '',
      '提出日': data.submissionDate || ''
    };

    var result = ConstructionRepository.save(constructionData);

    // 作業所注文協議書No.が空の場合、工事番号を自動セット
    if (!result['作業所注文協議書ＮＯ．'] && result['工事番号']) {
      result['作業所注文協議書ＮＯ．'] = result['工事番号'];
      ConstructionRepository.save(result);
    }

    Logger.log('工事保存完了: ' + result['工事番号']);
    return { success: true, id: result['工事番号'] };
  } catch (e) {
    Logger.log('工事保存エラー: ' + e.message);
    return { success: false, error: e.message };
  }
}

/**
 * 指定された工事を一括ロック（決算確定）
 * @param {string[]} constructionIds - ロック対象の工事番号リスト
 */
function lockConstructions(constructionIds) {
  Logger.log('lockConstructions開始 - 対象: ' + constructionIds.length + '件');
  try {
    var count = 0;
    for (var i = 0; i < constructionIds.length; i++) {
      var existing = ConstructionRepository.findById(constructionIds[i]);
      if (existing) {
        existing['ロック'] = 'TRUE';
        ConstructionRepository.save(existing);
        count++;
      }
    }
    Logger.log('ロック完了: ' + count + '件');
    return { success: true, lockedCount: count };
  } catch (e) {
    Logger.log('一括ロックエラー: ' + e.message);
    return { success: false, error: e.message };
  }
}

/**
 * 工事のロックを解除（管理者用）
 * @param {string} constructionId - ロック解除対象の工事番号
 */
function unlockConstruction(constructionId) {
  Logger.log('unlockConstruction開始 - 対象: ' + constructionId);
  try {
    var existing = ConstructionRepository.findById(constructionId);
    if (!existing) {
      return { success: false, error: '工事が見つかりません: ' + constructionId };
    }
    existing['ロック'] = '';
    ConstructionRepository.save(existing);
    Logger.log('ロック解除完了: ' + constructionId);
    return { success: true };
  } catch (e) {
    Logger.log('ロック解除エラー: ' + e.message);
    return { success: false, error: e.message };
  }
}

// ========================================
// 契約協議 API
// ========================================

/**
 * 契約協議一覧を取得
 * 発注者/契約者のIDは顧客名に変換して返す
 */
function getContracts(department) {
  Logger.log('getContracts開始 - department: ' + department);
  try {
    var contracts;
    var customerMap = MasterLookup.getCustomerMap();
    Logger.log('customerMap取得完了');

    contracts = ContractRepository.findAll();
    Logger.log('contracts取得: ' + contracts.length + '件');
    // 新しいデータを先頭に表示
    return contracts.map(function(row) {
      var customerId = row['発注者'];
      return {
        contractId: row['協議書ID'] || '',
        projectName: row['工事名'] || '',
        customerId: customerId || '',
        customerName: customerMap[customerId] || customerId || '',
        contractAmount: row['契約金額'] || 0,
        status: row['承認ステータス'] || '作成中',
        createdAt: formatDateForApi(row['作成日']),
        constructionCategory: row['工事区分'] || ''
      };
    }).reverse();
  } catch (e) {
    Logger.log('契約協議取得エラー: ' + e.message);
    Logger.log('スタックトレース: ' + e.stack);
    return [];
  }
}

/**
 * 日付をAPI返却用の文字列に変換
 * Date型やその他の形式を安全に文字列化
 */
function formatDateForApi(dateValue) {
  if (!dateValue) return '';
  if (dateValue instanceof Date) {
    return Utilities.formatDate(dateValue, 'Asia/Tokyo', 'yyyy-MM-dd');
  }
  return String(dateValue);
}

/**
 * オブジェクト/配列内のDate型を再帰的にyyyy-MM-dd文字列に変換
 * google.script.runでクライアントに返す前に必ず通す
 */
function sanitizeDates(obj) {
  if (obj === null || obj === undefined) return obj;
  if (obj instanceof Date) {
    return Utilities.formatDate(obj, 'Asia/Tokyo', 'yyyy-MM-dd');
  }
  if (Array.isArray(obj)) {
    return obj.map(function(item) { return sanitizeDates(item); });
  }
  if (typeof obj === 'object') {
    var result = {};
    Object.keys(obj).forEach(function(key) {
      result[key] = sanitizeDates(obj[key]);
    });
    return result;
  }
  return obj;
}

/**
 * 契約協議を1件取得
 * 発注者/契約者のIDは顧客名に変換して追加
 */
function getContractById(contractId, department) {
  Logger.log('getContractById開始 - contractId: ' + contractId + ', department: ' + department);
  try {
    var contract;
    var customerMap = MasterLookup.getCustomerMap();

    contract = ContractRepository.findById(contractId);
    Logger.log('findById結果: ' + (contract ? '見つかった' : 'null'));
    if (contract) {
      var customerId = contract['発注者'];
      contract['発注者_表示'] = customerMap[customerId] || customerId;
    }

    // 日付フィールドを文字列に変換（シリアライズ対策）
    if (contract) {
      contract = convertDatesToStrings(contract);
    }

    Logger.log('getContractById結果: ' + (contract ? 'データあり' : 'null'));
    return contract;
  } catch (e) {
    Logger.log('契約協議取得エラー: ' + e.message);
    return null;
  }
}

/**
 * オブジェクト内の日付を文字列に変換
 */
function convertDatesToStrings(obj) {
  var result = {};
  for (var key in obj) {
    if (obj.hasOwnProperty(key)) {
      var value = obj[key];
      if (value instanceof Date) {
        result[key] = Utilities.formatDate(value, 'Asia/Tokyo', 'yyyy-MM-dd');
      } else {
        result[key] = value;
      }
    }
  }
  return result;
}

/**
 * 契約協議を保存
 * 保存後、承認済の場合は工事管理に自動登録する
 * 区分「追加」の場合は元工事に紐付ける
 */
function saveContract(data, department) {
  try {
    var result;
    var contractId;

    result = ContractRepository.save(data);
    contractId = result['協議書ID'];

    var projectId = null;
    var isAdditional = data['区分'] === '追加' && data['関連元協議書ID'];

    // 承認済の場合、工事管理に自動登録
    Logger.log('saveContract - 承認ステータス: ' + data['承認ステータス'] + ', 工事担当者: "' + (data['工事担当者'] || '') + '"');
    if (data['承認ステータス'] === '承認済') {
      if (isAdditional) {
        // 追加工事: 元協議書の案件ID・工事番号を引き継ぐ
        projectId = registerAdditionalToExistingProject(data, contractId, department);
      } else {
        projectId = registerProjectFromContract(data, contractId, department);
      }
    }

    // 承認済かつ案件登録成功時、工事管理にも自動登録
    if (projectId) {
      if (isAdditional) {
        registerAdditionalToExistingConstruction(data, contractId, department, projectId);
      } else {
        registerConstructionFromContract(data, contractId, department, projectId);
      }
    }

    return { success: true, id: contractId, projectId: projectId };
  } catch (e) {
    Logger.log('契約協議保存エラー: ' + e.message);
    return { success: false, error: e.message };
  }
}

/**
 * 追加工事を元の案件に紐付ける
 * 元協議書IDから案件を検索し、契約金額を合算更新する
 */
function registerAdditionalToExistingProject(contractData, contractId, department) {
  try {
    var originalContractId = contractData['関連元協議書ID'];
    var existingProject = ProjectRepository.findByContractId(originalContractId);

    if (!existingProject) {
      Logger.log('元協議書の案件が見つかりません: ' + originalContractId + ' → 新規案件として登録');
      return registerProjectFromContract(contractData, contractId, department);
    }

    // 元案件の契約金額に追加分を合算
    var originalAmount = Number(existingProject['契約金額']) || 0;
    var additionalAmount = Number(contractData['契約金額']) || 0;
    existingProject['契約金額'] = originalAmount + additionalAmount;

    ProjectRepository.save(existingProject);
    Logger.log('追加工事を案件に合算: ' + existingProject['案件ID'] + ' (元: ' + originalContractId + ', 追加: ' + contractId + ', 合計金額: ' + existingProject['契約金額'] + ')');
    return existingProject['案件ID'];
  } catch (e) {
    Logger.log('追加工事の案件紐付けエラー: ' + e.message);
    return null;
  }
}

/**
 * 追加工事を元の工事番号に紐付ける
 * 元協議書IDから工事を検索し、同じ工事番号で関連付ける
 */
function registerAdditionalToExistingConstruction(contractData, contractId, department, projectId) {
  try {
    var originalContractId = contractData['関連元協議書ID'];
    var existingConstruction = ConstructionRepository.findByContractId(originalContractId);

    if (!existingConstruction) {
      Logger.log('元協議書の工事が見つかりません: ' + originalContractId + ' → 新規工事として登録');
      return registerConstructionFromContract(contractData, contractId, department, projectId);
    }

    Logger.log('追加工事を同一工事番号に紐付け: ' + existingConstruction['工事番号'] + ' (元: ' + originalContractId + ', 追加: ' + contractId + ')');
    // 同じ工事番号を使って工事管理に追加登録は不要（元工事に紐付けるだけ）
    // 関連契約協議書IDを追加分のIDで記録はしない（元工事の情報を維持）
  } catch (e) {
    Logger.log('追加工事の工事紐付けエラー: ' + e.message);
  }
}

/**
 * 契約協議書のデータを工事管理テーブルに登録
 * 承認済の契約協議書から工事を自動作成
 * @param {Object} contractData - 契約協議データ
 * @param {string} contractId - 契約協議書ID
 * @param {string} department - 部署
 * @returns {string|null} 登録された工事ID
 */
function registerProjectFromContract(contractData, contractId, department) {
  try {
    // 既存の工事を検索（重複防止）- 重複行がある場合は全件取得
    var allMatches = findData('案件管理', { '関連契約協議書ID': contractId });
    var existing = allMatches.length > 0 ? allMatches[0] : null;

    // 重複行がある場合は2行目以降を削除
    if (allMatches.length > 1) {
      Logger.log('重複案件を検出: ' + allMatches.length + '件 → 2行目以降を削除');
      for (var i = 1; i < allMatches.length; i++) {
        var dupId = allMatches[i]['案件ID'];
        if (dupId) {
          try {
            deleteData('案件管理', '案件ID', dupId);
            Logger.log('重複案件を削除: ' + dupId);
          } catch (delErr) {
            Logger.log('重複案件削除エラー: ' + delErr.message);
          }
        }
      }
    }

    // フロントエンドから渡された工事担当者を取得
    var managerName = contractData['工事担当者'] || '';
    Logger.log('registerProjectFromContract - 工事担当者: "' + managerName + '", 既存案件: ' + (existing ? existing['案件ID'] : 'なし'));

    var projectData = {
      '関連契約協議書ID': contractId,
      '案件名': contractData['工事名'] || '',
      '顧客名': contractData['発注者'] || '',
      '契約金額': contractData['契約金額'] || 0,
      '担当者名': managerName,
      '契約日': contractData['作成日'] || '',
      '着工日': contractData['工期_開始'] || '',
      '竣工予定日': contractData['工期_終了'] || '',
      '案件ステータス': '契約済',
      '部署': 'ソリューション事業部',
      '振り分けカテゴリー': 'ソリューション',
      '削除フラグ': '',
      '削除日時': '',
      '削除者': ''
    };

    // 既存の工事がある場合は工事IDを引き継ぐ（更新）、新規の場合は登録日を設定
    if (existing) {
      projectData['案件ID'] = existing['案件ID'];
    } else {
      projectData['登録日'] = Utilities.formatDate(new Date(), 'Asia/Tokyo', 'yyyy-MM-dd');
    }

    var savedProject = ProjectRepository.save(projectData);

    // 書き込みを強制フラッシュ
    SpreadsheetApp.flush();

    // 保存後にデータを読み戻して検証
    var verify = ProjectRepository.findById(savedProject['案件ID']);
    Logger.log('工事管理に登録完了: ' + savedProject['案件ID'] + ', 担当者名: "' + managerName + '" (契約協議書: ' + contractId + ')');
    Logger.log('検証 - 案件管理シート読み戻し: 担当者名="' + (verify ? verify['担当者名'] : 'NOT FOUND') + '"');

    return savedProject['案件ID'];
  } catch (e) {
    // 工事登録エラーは契約協議保存自体には影響させない
    Logger.log('工事管理への登録エラー（契約協議保存は成功）: ' + e.message);
    return null;
  }
}

/**
 * 承認済契約協議書一覧を取得（工事登録用）
 * @param {string} department - 部署（省略時は両方）
 * @returns {Array} 承認済契約協議書一覧
 */
function getApprovedContractsForProject(department) {
  try {
    var result = [];
    var customerMap = MasterLookup.getCustomerMap();

    var contracts = ContractRepository.findAll();
    contracts.forEach(function(row) {
      if (row['承認ステータス'] === '承認済') {
        var customerId = row['発注者'];
        result.push({
          contractId: row['協議書ID'],
          projectName: row['工事名'],
          customerId: customerId,
          customerName: customerMap[customerId] || customerId,
          contractAmount: row['契約金額'] || 0,
          department: 'ソリューション事業部',
          createdAt: formatDateForApi(row['作成日'])
        });
      }
    });

    return result;
  } catch (e) {
    Logger.log('承認済契約協議取得エラー: ' + e.message);
    return [];
  }
}

/**
 * 契約協議書のデータを工事管理テーブルに登録
 * 既存の工事がある場合は更新、なければ新規作成
 * @param {Object} contractData - 契約協議データ
 * @param {string} contractId - 契約協議書ID
 * @param {string} department - 部署
 * @param {string} projectId - 関連工事ID（承認済の場合に設定）
 */
function registerConstructionFromContract(contractData, contractId, department, projectId) {
  try {
    // 既存の工事を検索 - 重複がある場合は2行目以降を削除
    var allConstMatches = findData('工事管理', { '関連契約協議書ID': contractId });
    var existing = allConstMatches.length > 0 ? allConstMatches[0] : null;
    if (allConstMatches.length > 1) {
      Logger.log('重複工事を検出: ' + allConstMatches.length + '件 → 2行目以降を削除');
      for (var ci = 1; ci < allConstMatches.length; ci++) {
        var dupNo = allConstMatches[ci]['工事番号'];
        if (dupNo) {
          try { deleteData('工事管理', '工事番号', dupNo); } catch (e2) { Logger.log('重複工事削除エラー: ' + e2.message); }
        }
      }
    }

    var location = contractData['工事場所'] || '';

    var startDate = contractData['工期_開始'] || '';
    var endDate = contractData['工期_終了'] || '';

    // 契約協議書のカラム名から工事管理のカラム名にマッピング
    var constructionData = {
      '関連契約協議書ID': contractId,
      '関連案件ID': projectId || (existing ? existing['関連案件ID'] : ''),
      '工事名': contractData['工事名'] || '',
      '工事場所': location,
      '工事担当社員': contractData['工事担当者'] || contractData['作成者'] || '',
      '営業担当社員': contractData['営業担当者'] || '',
      '工期始': startDate,
      '工期終': endDate,
      '進捗ステータス': existing ? existing['進捗ステータス'] : '未着工',
      '部署': department,
      '削除フラグ': '',
      '削除日時': '',
      '削除者': ''
    };

    // 既存の工事がある場合は工事番号を引き継ぐ（更新）
    if (existing) {
      constructionData['工事番号'] = existing['工事番号'];
    }

    var savedConstruction = ConstructionRepository.save(constructionData);

    // 作業所注文協議書No.が空の場合、工事番号を自動セット
    if (!savedConstruction['作業所注文協議書ＮＯ．'] && savedConstruction['工事番号']) {
      savedConstruction['作業所注文協議書ＮＯ．'] = savedConstruction['工事番号'];
      ConstructionRepository.save(savedConstruction);
    }

    Logger.log('工事管理に登録完了: ' + savedConstruction['工事番号'] + ' (契約協議書: ' + contractId + ', 工事: ' + (projectId || '未設定') + ')');
  } catch (e) {
    // 工事登録エラーは契約協議保存自体には影響させない
    Logger.log('工事管理への登録エラー（契約協議保存は成功）: ' + e.message);
  }
}

/**
 * 契約協議書をExcel出力
 * テンプレートの種類を自動判別（スプレッドシート/ドキュメント）
 * @param {string} contractId - 契約協議書ID
 * @param {string} department - 部署
 * @returns {Object} 結果オブジェクト {success, url, filename, error}
 */
function exportContractToExcel(contractId, department) {
  try {
    // 1. テンプレートIDの確認
    var templateId = OUTPUT_CONFIG.TEMPLATE_CONTRACT_SOLUTION;

    if (!templateId) {
      return {
        success: false,
        error: 'テンプレートファイルIDが設定されていません。OUTPUT_CONFIGを確認してください。'
      };
    }

    // 2. 契約協議データを取得
    var contract = getContractById(contractId, department);
    if (!contract) {
      return { success: false, error: '契約協議が見つかりません: ' + contractId };
    }

    // 3. 関連データを取得してマージ（AppSheetリレーション参照対応）
    var enrichedData = enrichContractDataForDocument(contract, department);

    // 4. テンプレートの種類を判別
    var templateFile = DriveApp.getFileById(templateId);
    var mimeType = templateFile.getMimeType();

    // 5. ファイル名を生成
    var projectName = contract['工事名'] || contractId;
    var fileName = generateFileName('【契約】', projectName, 'xlsx');

    var downloadUrl;
    var fileId;

    // Googleドキュメントの場合（Excel出力不可、PDFで代替）
    if (mimeType === MimeType.GOOGLE_DOCS) {
      var document = copyDocumentAndFill(
        templateId,
        enrichedData,
        null,
        fileName.replace('.xlsx', ''),
        OUTPUT_CONFIG.OUTPUT_FOLDER_ID
      );
      // ドキュメントのダウンロードURL
      downloadUrl = 'https://docs.google.com/document/d/' + document.getId() + '/export?format=docx';
      fileId = document.getId();
      fileName = fileName.replace('.xlsx', '.docx');
    }
    // Googleスプレッドシートの場合
    else {
      var spreadsheet = copyTemplateAndFillAuto(
        templateId,
        enrichedData,
        fileName.replace('.xlsx', ''),
        OUTPUT_CONFIG.OUTPUT_FOLDER_ID
      );
      downloadUrl = getExcelDownloadUrl(spreadsheet);
      fileId = spreadsheet.getId();
    }

    return {
      success: true,
      url: downloadUrl,
      filename: fileName,
      fileId: fileId
    };
  } catch (e) {
    Logger.log('契約協議書Excel出力エラー: ' + e.message);
    return { success: false, error: e.message };
  }
}

/**
 * 契約協議書（小口用）をExcel出力
 * 小口契約協議テンプレートにデータを埋め込んでスプレッドシートを生成
 * @param {string} contractId - 契約協議書ID
 * @param {string} department - 部署
 * @returns {Object} 結果オブジェクト {success, url, filename, fileId, error}
 */
function exportContractSmallToExcel(contractId, department) {
  try {
    var templateId = OUTPUT_CONFIG.TEMPLATE_CONTRACT_SMALL;
    if (!templateId) {
      return { success: false, error: '小口用テンプレートファイルIDが設定されていません。' };
    }

    var contract = getContractById(contractId, department);
    if (!contract) {
      return { success: false, error: '契約協議が見つかりません: ' + contractId };
    }

    // 関連データを取得してマージ + 小口用の追加フィールドを生成
    var enrichedData = enrichContractDataForDocument(contract, department);
    enrichSmallContractData(enrichedData, contract);

    var projectName = contract['工事名'] || contractId;
    var fileName = generateFileName('【小口契約】', projectName, 'xlsx');

    // スプレッドシートテンプレートをコピーしてプレースホルダー置換
    var spreadsheet = copyTemplateAndFillAuto(
      templateId,
      enrichedData,
      fileName.replace('.xlsx', ''),
      OUTPUT_CONFIG.OUTPUT_FOLDER_ID
    );

    // 工事施工協議書シート: 工事場所の複写を修正
    // テンプレートでは D13:H15, I13:I15, J13:N15 の3つの結合セルが
    // 小口契約協議シートの個別セルを参照しているが、小口契約協議側は
    // D13:AI15 の1つの結合セルのため I13, J13 が空になる。
    // → 3つの結合セルを1つに統合して住所全文を直接書き込む
    var sekouSheet = spreadsheet.getSheetByName('工事施工協議書');
    if (sekouSheet) {
      sekouSheet.getRange('D13:N15').breakApart();
      sekouSheet.getRange('D13:N15').merge();
      sekouSheet.getRange('D13').setValue(enrichedData['工事場所_府県'] || '');
    }
    SpreadsheetApp.flush();

    return {
      success: true,
      url: getExcelDownloadUrl(spreadsheet),
      filename: fileName,
      fileId: spreadsheet.getId()
    };
  } catch (e) {
    Logger.log('小口契約協議書Excel出力エラー: ' + e.message);
    return { success: false, error: e.message };
  }
}

/**
 * 小口契約協議書用の追加データを生成
 * 住所分割、日付分割、入金方法テキスト生成など
 * @param {Object} data - enrichContractDataForDocumentの結果（この関数で直接追加される）
 * @param {Object} contract - 元の契約データ
 */
function enrichSmallContractData(data, contract) {
  // === 工事場所（テンプレートが1セル結合のため分割不要）===
  data['工事場所_府県'] = String(contract['工事場所'] || '');

  // === 契約協議書日付の年月日分割 ===
  var contractDate = String(contract['契約協議書日付'] || '');
  var dateParts = contractDate.match(/(\d{4})[\/\-](\d{1,2})[\/\-](\d{1,2})/);
  if (dateParts) {
    data['契約協議書日付_年'] = dateParts[1];
    data['契約協議書日付_月'] = dateParts[2];
    data['契約協議書日付_日'] = dateParts[3];
  } else {
    data['契約協議書日付_年'] = '';
    data['契約協議書日付_月'] = '';
    data['契約協議書日付_日'] = '';
  }

  // === 工期の年月日分割 ===
  var periodStart = String(contract['工期_開始'] || '');
  var startParts = periodStart.match(/(\d{4})[\/\-](\d{1,2})[\/\-](\d{1,2})/);
  if (startParts) {
    data['工期開始_年'] = startParts[1];
    data['工期開始_月'] = startParts[2];
    data['工期開始_日'] = startParts[3];
  } else {
    data['工期開始_年'] = '';
    data['工期開始_月'] = '';
    data['工期開始_日'] = '';
  }

  var periodEnd = String(contract['工期_終了'] || '');
  var endParts = periodEnd.match(/(\d{4})[\/\-](\d{1,2})[\/\-](\d{1,2})/);
  if (endParts) {
    data['工期終了_年'] = endParts[1];
    data['工期終了_月'] = endParts[2];
    data['工期終了_日'] = endParts[3];
  } else {
    data['工期終了_年'] = '';
    data['工期終了_月'] = '';
    data['工期終了_日'] = '';
  }

  // === 金額決定日の月日分割 ===
  var priceDate = String(contract['金額決定日'] || '');
  var priceParts = priceDate.match(/(\d{4})[\/\-](\d{1,2})[\/\-](\d{1,2})/);
  if (priceParts) {
    data['金額決定_月'] = priceParts[2];
    data['金額決定_日'] = priceParts[3];
  } else {
    data['金額決定_月'] = '';
    data['金額決定_日'] = '';
  }

  // === 契約書 要/不要 （選択された方だけ表示）===
  var contractRequired = String(contract['契約書要不要'] || '');
  data['契約書_要'] = contractRequired === '要' ? '要' : '';
  data['契約書_不要'] = contractRequired === '不要' ? '不要' : '';

  // === 消費税・契約金額（数式が壊れるのでコードで計算）===
  var constructionPrice = Number(contract['工事価格']) || 0;
  data['消費税額'] = Math.floor(constructionPrice * 0.1);
  data['契約金額_税込'] = constructionPrice + data['消費税額'];

  // === 粗利率%（数式が壊れるのでコードで計算）===
  var submittedEstimate = Number(contract['提出見積金額']) || 0;
  var grossProfit = Number(contract['NET_粗付加金額']) || 0;
  if (submittedEstimate > 0) {
    data['粗付加金額_率'] = (grossProfit / submittedEstimate * 100).toFixed(1);
  } else {
    data['粗付加金額_率'] = '';
  }

  // === 精算カテゴリーのテキスト化 ===
  var settlement = String(contract['精算カテゴリー'] || '');
  if (settlement === '契約') {
    data['精算カテゴリー'] = '（●契約）（精算）（未定）';
  } else if (settlement === '精算') {
    data['精算カテゴリー'] = '（契約）（●精算）（未定）';
  } else if (settlement === '未定') {
    data['精算カテゴリー'] = '（契約）（精算）（●未定）';
  } else {
    data['精算カテゴリー'] = '（契約）（精算）（未定）';
  }

  // === 入金方法テキスト ===
  // 入金方法_現金/手形はスプレッドシートのセル書式によって以下のいずれかで取得される:
  //   - パーセント書式: 数値 1 (=100%), 0.7 (=70%) など
  //   - 文字列: "100%", "70%" など
  //   - 数値（%なし）: 100, 70 など
  var cashVal = contract['入金方法_現金'];
  var noteVal = contract['入金方法_手形'];
  var site = contract['入金方法_サイト'] || '';

  // パーセント数値を正規化（表示用の整数%に変換）
  function normalizePercent(val) {
    if (val === null || val === undefined || val === '') return 0;
    var str = String(val).replace(/%/g, '').trim();
    var num = parseFloat(str);
    if (isNaN(num)) return 0;
    // 0〜1の範囲なら小数パーセント（例: 0.7 = 70%, 1 = 100%）
    if (num > 0 && num <= 1) {
      num = Math.round(num * 100);
    }
    return num;
  }

  var cashPct = normalizePercent(cashVal);
  var notePct = normalizePercent(noteVal);
  var siteText = site ? site + '日' : '';

  // 手形が0%の場合は現金のみ表示
  var lines = ['現金　' + cashPct + '%'];
  if (notePct > 0) {
    lines.push('手形　' + notePct + '%' + (siteText ? '　サイト　' + siteText : ''));
  }
  data['入金方法_テキスト'] = lines.join('\n');

  // === 資金調達方法テキスト ===
  var funding = String(contract['資金調達方法'] || '');
  var loanDetail = String(contract['資金調達方法_融資'] || '');
  if (funding === '自己資金') {
    data['資金調達方法_テキスト'] = '　・●自己資金\n　・融資（　　　　　　　　）・その他';
  } else if (funding === '融資') {
    data['資金調達方法_テキスト'] = '　・自己資金\n　・●融資（' + loanDetail + '）・その他';
  } else {
    data['資金調達方法_テキスト'] = '　・自己資金\n　・融資（　　　　　　　　）・その他';
  }

  // === 工事番号（工事管理から取得を試みる）===
  // 契約協議書IDに紐づく工事番号があれば取得
  try {
    var constructions = ConstructionRepository.findAll();
    for (var i = 0; i < constructions.length; i++) {
      if (constructions[i]['関連契約協議書ID'] === contract['協議書ID']) {
        data['工事番号'] = constructions[i]['工事番号'] || '';
        break;
      }
    }
  } catch (e) {
    Logger.log('工事番号取得エラー: ' + e.message);
  }
  if (!data['工事番号']) {
    data['工事番号'] = '';
  }
}

/**
 * 契約協議書をPDF出力
 * テンプレートの種類を自動判別（スプレッドシート/ドキュメント）
 * @param {string} contractId - 契約協議書ID
 * @param {string} department - 部署
 * @returns {Object} 結果オブジェクト {success, url, filename, error}
 */
function exportContractToPdf(contractId, department) {
  try {
    // 1. テンプレートIDの確認
    var templateId = OUTPUT_CONFIG.TEMPLATE_CONTRACT_SOLUTION;

    if (!templateId) {
      return {
        success: false,
        error: 'テンプレートファイルIDが設定されていません。OUTPUT_CONFIGを確認してください。'
      };
    }

    // 2. 契約協議データを取得
    var contract = getContractById(contractId, department);
    if (!contract) {
      return { success: false, error: '契約協議が見つかりません: ' + contractId };
    }

    // 3. 関連データを取得してマージ
    var enrichedData = enrichContractDataForDocument(contract, department);

    // 4. テンプレートの種類を判別
    var templateFile = DriveApp.getFileById(templateId);
    var mimeType = templateFile.getMimeType();

    // 5. ファイル名を生成
    var projectName = contract['工事名'] || contractId;
    var pdfFileName = generateFileName('【契約】', projectName, 'pdf');

    var downloadUrl;
    var fileId;

    // Googleドキュメントの場合
    if (mimeType === MimeType.GOOGLE_DOCS) {
      var document = copyDocumentAndFill(
        templateId,
        enrichedData,
        null,
        pdfFileName.replace('.pdf', ''),
        OUTPUT_CONFIG.OUTPUT_FOLDER_ID
      );
      var pdfFile = exportDocumentAsPdf(document, pdfFileName, OUTPUT_CONFIG.OUTPUT_FOLDER_ID);
      downloadUrl = getFileDownloadUrl(pdfFile);
      fileId = pdfFile.getId();
    }
    // Googleスプレッドシートの場合
    else {
      var spreadsheet = copyTemplateAndFillAuto(
        templateId,
        enrichedData,
        pdfFileName.replace('.pdf', ''),
        OUTPUT_CONFIG.OUTPUT_FOLDER_ID
      );
      var pdfFile = exportSpreadsheetAsPdf(spreadsheet, pdfFileName, OUTPUT_CONFIG.OUTPUT_FOLDER_ID);
      downloadUrl = getFileDownloadUrl(pdfFile);
      fileId = pdfFile.getId();
    }

    return {
      success: true,
      url: downloadUrl,
      filename: pdfFileName,
      fileId: fileId
    };
  } catch (e) {
    Logger.log('契約協議書PDF出力エラー: ' + e.message);
    return { success: false, error: e.message };
  }
}

// ========================================
// 発注管理 API
// ========================================

/**
 * 発注一覧を取得（全件）
 * スプレッドシートのカラム名をそのまま返す（日本語カラム名）
 * 仕入先コードは協力会社名_表示フィールドに名前変換して追加
 */
function getOrders() {
  try {
    var ss = getSpreadsheet();
    var sheet = ss.getSheetByName('発注業者管理');

    if (!sheet || sheet.getLastRow() < 2) {
      return [];
    }

    var lastRow = sheet.getLastRow();
    var lastCol = sheet.getLastColumn();
    var headers = sheet.getRange(1, 1, 1, lastCol).getValues()[0];
    var data = sheet.getRange(2, 1, lastRow - 1, lastCol).getValues();

    // マスタルックアップを取得（キャッシュ利用）
    var vendorMap = MasterLookup.getVendorMap();

    var orders = data.map(function(row) {
      var obj = {};
      headers.forEach(function(header, index) {
        var value = row[index];
        // Date型は文字列に変換（シリアライズ問題対策）
        if (value instanceof Date) {
          value = Utilities.formatDate(value, Session.getScriptTimeZone(), 'yyyy-MM-dd');
        }
        obj[header] = value;
      });

      // ID→名前変換を付与（「施主名(注文決定業者名)」列には実際は仕入先コードが格納されている）
      var vendorId = obj['施主名(注文決定業者名)'];
      obj['仕入先コード'] = vendorId;
      obj['協力会社名_表示'] = vendorMap[vendorId] || vendorId;

      return obj;
    });

    // 新しいデータを先頭に表示
    orders.reverse();

    return orders;
  } catch (e) {
    Logger.log('発注取得エラー: ' + e.message);
    return [];
  }
}

/**
 * 発注を1件取得（発注IDで検索）
 * 全カラムを日本語キーで返す（OrderFormの編集用）
 */
function getOrderById(orderId) {
  try {
    var order = OrderRepository.findById(orderId);
    if (!order) {
      Logger.log('発注が見つかりません: ' + orderId);
      return null;
    }

    // Date型を文字列に変換（シリアライズ問題対策）
    var result = {};
    Object.keys(order).forEach(function(key) {
      var value = order[key];
      if (value instanceof Date) {
        value = Utilities.formatDate(value, Session.getScriptTimeZone(), 'yyyy-MM-dd');
      }
      result[key] = value;
    });

    // マスタルックアップを取得
    var vendorMap = MasterLookup.getVendorMap();
    var vendorId = result['施主名(注文決定業者名)'];
    result['仕入先コード'] = vendorId;
    result['協力会社名_表示'] = vendorMap[vendorId] || vendorId;

    Logger.log('getOrderById成功: ' + orderId);
    return result;
  } catch (e) {
    Logger.log('発注取得エラー: ' + e.message);
    return null;
  }
}

/**
 * 発注一覧を取得（工事番号で絞り込み）
 */
function getOrdersByConstructionId(constructionId) {
  try {
    var orders = OrderRepository.findByConstructionId(constructionId);
    var vendorMap = MasterLookup.getVendorMap();

    return orders.map(function(row) {
      var vendorId = row['施主名(注文決定業者名)'];
      return {
        orderId: row['発注ID'],
        constructionId: row['関連工事番号'],
        vendorId: vendorId,
        vendorName: vendorMap[vendorId] || vendorId,
        orderContent: row['注文内容'],
        orderAmount: Number(row['注文金額'] || 0),
        budgetAmount: Number(row['実行予算金額'] || 0),
        paymentStatus: row['支払ステータス(協力会社への)'],
        vendorRank: row['注文決定業者のランク']
      };
    }).reverse();
  } catch (e) {
    Logger.log('発注取得エラー: ' + e.message);
    return [];
  }
}

/**
 * 発注を保存（24カラム対応）
 * フロントエンドから日本語カラム名のデータをそのまま受け取る
 */
function saveOrder(data) {
  try {
    // フロントエンドから日本語カラム名でデータが来る想定
    // OrderFormのtoSheetData()で変換済み
    var sheetData = {
      '発注ID': data['発注ID'],
      '関連工事番号': data['関連工事番号'],
      '施主名(注文決定業者名)': data['施主名(注文決定業者名)'],
      '注文内容': data['注文内容'],
      '発注区分': data['発注区分'] || 'その他',
      '注文金額': data['注文金額'],
      '支払条件(毎月◯日締め切り)': data['支払条件(毎月◯日締め切り)'],
      '支払条件(翌月◯日支払)': data['支払条件(翌月◯日支払)'],
      '支払条件(出来高◯%以内)': data['支払条件(出来高◯%以内)'],
      '支払条件(現金%)': data['支払条件(現金%)'],
      '支払条件(手形%)': data['支払条件(手形%)'],
      '支払ステータス(協力会社への)': data['支払ステータス(協力会社への)'],
      '安推協特別会費': data['安推協特別会費'],
      '協議内容': data['協議内容'],
      '注文条件': data['注文条件'],
      'かし担保': data['かし担保'],
      '備考': data['備考'],
      '実行予算金額': data['実行予算金額'],
      '注文決定業者のランク': data['注文決定業者のランク'],
      '注文書発行日': data['注文書発行日'],
      '注文書番号': data['注文書番号'],
      '建設業の許可': data['建設業の許可'],
      '注文協議書パス': data['注文協議書パス'],
      '注文書パス': data['注文書パス'],
      '返送された注文請書': data['返送された注文請書']
    };

    var result = OrderRepository.save(sheetData);
    return { success: true, id: result['発注ID'] };
  } catch (e) {
    Logger.log('発注保存エラー: ' + e.message);
    return { success: false, error: e.message };
  }
}

/**
 * 発注登録時に支払データを自動作成
 * 発注情報と協力会社マスタの支払条件から支払レコードを生成
 * @param {string} orderId - 発注ID
 * @returns {Object} 作成結果
 */
function createPaymentFromOrder(orderId) {
  try {
    var order = OrderRepository.findById(orderId);
    if (!order) {
      return { success: false, error: '発注が見つかりません: ' + orderId };
    }

    var vendorId = order['施主名(注文決定業者名)'];
    var orderAmount = Number(order['注文金額'] || 0);
    var orderCategory = order['発注区分'] || 'その他';

    if (!vendorId || orderAmount <= 0) {
      return { success: false, error: '協力会社または注文金額が未設定です' };
    }

    // 支払条件を適用（でんさい・安推協会費・振込の計算）
    // まずマスタのデフォルト値を取得
    var termsResult = applyVendorPaymentTerms(vendorId, orderAmount, orderCategory);
    if (!termsResult.success) {
      return termsResult;
    }
    var terms = termsResult.data;

    // 注文書にカスタム支払い条件が設定されている場合、マスタの値をオーバーライドする
    var orderNoteRate = order['支払条件(手形%)'];
    var orderCashRate = order['支払条件(現金%)'];
    if (orderNoteRate !== undefined && orderNoteRate !== '' && orderCashRate !== undefined && orderCashRate !== '') {
      var customNoteRatio = Number(orderNoteRate) || 0;
      var amountExcludingTax = orderAmount;
      var tax = Math.floor(amountExcludingTax * 0.1);
      var amountIncludingTax = amountExcludingTax + tax;
      // 手形額: 税抜金額 × カスタム比率（万円未満切捨て）
      terms.notePayment = Math.floor(amountExcludingTax * customNoteRatio / 100 / 10000) * 10000;
      // 振込金額: 税込金額 - 手形額 - 安推協会費（帳尻合わせ）
      terms.cashPayment = amountIncludingTax - terms.notePayment - terms.safetyFee;
      terms.noteRatio = customNoteRatio;
    }

    // 注文書にカスタム締日・支払日が設定されている場合もオーバーライド
    var orderClosingDay = order['支払条件(毎月◯日締め切り)'];
    var orderDueDay = order['支払条件(翌月◯日支払)'];
    if (orderClosingDay !== undefined && orderClosingDay !== '') {
      terms.closingDay = orderClosingDay;
    }
    if (orderDueDay !== undefined && orderDueDay !== '') {
      terms.paymentDay = orderDueDay;
    }

    // 支払予定日を計算
    var closingDay = Number(terms.closingDay) || 25;
    var paymentMonthOffset = Number(terms.paymentMonth) || 1;
    var paymentDay = terms.paymentDay;

    var now = new Date();
    var dueDate = new Date(now.getFullYear(), now.getMonth() + paymentMonthOffset, 1);
    if (paymentDay === '末') {
      dueDate = new Date(dueDate.getFullYear(), dueDate.getMonth() + 1, 0);
    } else {
      dueDate.setDate(Number(paymentDay) || 10);
    }
    var dueDateStr = Utilities.formatDate(dueDate, 'Asia/Tokyo', 'yyyy-MM-dd');
    var invoiceMonth = Utilities.formatDate(now, 'Asia/Tokyo', 'yyyy-MM');

    var paymentData = {
      orderId: orderId,
      projectId: order['関連工事番号'] || '',
      vendorId: vendorId,
      vendorName: terms.vendorName,
      invoiceMonth: invoiceMonth,
      closingDay: terms.closingDay,
      dueDate: dueDateStr,
      amountExcludingTax: orderAmount,
      orderCategory: orderCategory,
      safetyFee: terms.safetyFee,
      cashPayment: terms.cashPayment,
      notePayment: terms.notePayment,
      noteSite: terms.noteSite,
      settlementType: '通常',
      safetyFeeExempt: false,
      paymentStatus: '未払'
    };

    return savePayment(paymentData);
  } catch (e) {
    Logger.log('発注連動支払作成エラー: ' + e.message);
    return { success: false, error: e.message };
  }
}

/**
 * 契約協議承認時に顧客請求データを自動作成
 * @param {string} contractId - 協議書ID
 * @param {string} projectId - 案件ID（承認時に作成/取得された案件ID）
 * @returns {Object} 作成結果
 */
function createInvoiceFromContract(contractId, projectId) {
  try {
    var contract = ContractRepository.findById(contractId);
    if (!contract) {
      return { success: false, error: '契約協議が見つかりません: ' + contractId };
    }

    // 既に同じ契約から請求が作成されていないか確認（重複防止）
    var existingInvoices = InvoiceRepository.findByProjectId(projectId);
    var alreadyCreated = existingInvoices.some(function(inv) {
      return inv['備考'] && inv['備考'].indexOf('契約協議連動:' + contractId) >= 0;
    });
    if (alreadyCreated) {
      Logger.log('既に請求データが作成済み: ' + contractId);
      return { success: true, id: null, message: '既に作成済み' };
    }

    var contractAmount = Number(contract['契約金額']) || 0;
    if (contractAmount <= 0) {
      return { success: false, error: '契約金額が未設定です' };
    }

    // 契約協議書では顧客IDは「発注者」カラムに格納されている
    var customerId = contract['発注者'] || '';
    var customerMap = MasterLookup.getCustomerMap();
    var customerName = customerMap[customerId] || customerId || '';

    var amountExcludingTax = Number(contract['工事価格']) || 0;

    var invoiceData = {
      projectId: projectId,
      customerId: customerId,
      customerName: customerName,
      invoiceType: '',
      invoiceDate: '',
      amountExcludingTax: amountExcludingTax,
      dueDate: '',
      paymentStatus: '未入金',
      paymentAmount: 0,
      notes: '契約協議連動:' + contractId
    };

    return saveInvoice(invoiceData);
  } catch (e) {
    Logger.log('契約連動請求作成エラー: ' + e.message);
    return { success: false, error: e.message };
  }
}

/**
 * 見積もり詳細を取得（相見積もり比較用）
 * フロントエンドが日本語キーでアクセスするため、日本語キーで返す
 */
function getQuotesByOrderId(orderId) {
  try {
    Logger.log('getQuotesByOrderId開始 - orderId: ' + orderId);
    var quotes = QuoteRepository.findByOrderId(orderId);
    Logger.log('取得した見積もりデータ件数: ' + quotes.length);

    // フロントエンドは日本語キーでアクセスするため、そのまま返す
    return quotes.map(function(row) {
      return {
        '見積もり詳細ID': row['見積もり詳細ID'],
        '発注ID': row['発注ID'],
        '工事番号': row['工事番号'],
        '見積業者名': row['見積業者名'] || '',
        'TEL': row['TEL'] || '',
        '担当者': row['担当者'] || '',
        '見積金額': row['見積金額'] || 0,
        '当社NET': row['当社NET'] || 0,
        '最終協議金額': row['最終協議金額'] || 0,
        '選定結果': row['選定結果'] || '',
        '備考': row['備考'] || ''
      };
    });
  } catch (e) {
    Logger.log('見積もり取得エラー: ' + e.message);
    return [];
  }
}

/**
 * 見積もり詳細を保存（新規/更新）
 * @param {Object} data - 見積もりデータ
 * @returns {Object} {success, id, error}
 */
function saveQuote(data) {
  try {
    // フロントエンドのフィールド名をシートのカラム名に変換
    var sheetData = {
      '見積もり詳細ID': data.quoteId || data['見積もり詳細ID'] || '',
      '工事番号': data.constructionId || data['工事番号'] || '',
      '発注ID': data.orderId || data['発注ID'] || '',
      '見積業者名': data.vendorName || data['見積業者名'] || '',
      'TEL': data.phone || data['TEL'] || '',
      '担当者': data.contact || data['担当者'] || '',
      '見積金額': data.quoteAmount || data['見積金額'] || 0,
      '当社NET': data.netAmount || data['当社NET'] || 0,
      '最終協議金額': data.finalAmount || data['最終協議金額'] || 0
    };

    Logger.log('[saveQuote] 保存データ: ID=' + sheetData['見積もり詳細ID'] + ', 発注ID=' + sheetData['発注ID'] + ', 業者=' + sheetData['見積業者名']);
    var result = QuoteRepository.save(sheetData);
    Logger.log('[saveQuote] 保存結果: ID=' + result['見積もり詳細ID']);
    return { success: true, id: result['見積もり詳細ID'] };
  } catch (e) {
    Logger.log('見積もり保存エラー: ' + e.message);
    return { success: false, error: e.message };
  }
}

/**
 * 見積もり詳細を削除
 * @param {string} quoteId - 見積もり詳細ID
 * @returns {Object} {success, error}
 */
function deleteQuote(quoteId) {
  try {
    if (!quoteId) {
      return { success: false, error: '見積もり詳細IDが指定されていません' };
    }

    QuoteRepository.delete(quoteId);
    Logger.log('見積もり削除成功: ' + quoteId);
    return { success: true, id: quoteId };
  } catch (e) {
    Logger.log('見積もり削除エラー: ' + e.message);
    return { success: false, error: e.message };
  }
}

/**
 * 注文協議書を出力（Excel/PDF）
 * テンプレートの種類を自動判別（スプレッドシート/ドキュメント）
 * @param {string} orderId - 発注ID
 * @param {string} format - 出力形式（'excel' または 'pdf'）デフォルトはexcel
 * @returns {Object} 結果オブジェクト {success, url, filename, error}
 */
function exportOrderDocument(orderId, format) {
  try {
    format = format || 'excel';

    // 1. テンプレートIDの確認
    var templateId = OUTPUT_CONFIG.TEMPLATE_ORDER_DOCUMENT;
    if (!templateId) {
      return {
        success: false,
        error: '注文協議書テンプレートファイルIDが設定されていません。OUTPUT_CONFIGを確認してください。'
      };
    }

    // 2. 発注データを取得
    var order = OrderRepository.findById(orderId);
    if (!order) {
      return { success: false, error: '発注データが見つかりません: ' + orderId };
    }

    // 3. 関連データを取得してマージ（AppSheetリレーション参照対応）
    var enrichedData = enrichOrderDataForDocument(order);

    // 4. テンプレートの種類を判別
    var templateFile = DriveApp.getFileById(templateId);
    var mimeType = templateFile.getMimeType();

    // 5. ファイル名を生成
    var constructionId = order['関連工事番号'] || orderId;
    var projectName = enrichedData['関連工事番号.工事名'] || constructionId;
    var extension = format === 'pdf' ? 'pdf' : 'xlsx';
    var fileName = generateFileName('【注文協議】', projectName, extension);

    var downloadUrl;
    var fileId;

    // Googleドキュメントの場合
    if (mimeType === MimeType.GOOGLE_DOCS) {
      var document = copyDocumentAndFill(
        templateId,
        enrichedData,
        null,
        fileName.replace('.' + extension, ''),
        OUTPUT_CONFIG.OUTPUT_FOLDER_ID
      );

      if (format === 'pdf') {
        var pdfFile = exportDocumentAsPdf(document, fileName, OUTPUT_CONFIG.OUTPUT_FOLDER_ID);
        downloadUrl = getFileDownloadUrl(pdfFile);
        fileId = pdfFile.getId();
      } else {
        // ドキュメントはExcel出力不可、PDFで出力
        var pdfFile = exportDocumentAsPdf(document, fileName.replace('.xlsx', '.pdf'), OUTPUT_CONFIG.OUTPUT_FOLDER_ID);
        downloadUrl = getFileDownloadUrl(pdfFile);
        fileId = pdfFile.getId();
        fileName = fileName.replace('.xlsx', '.pdf');
      }
    }
    // Googleスプレッドシートの場合
    else {
      var spreadsheet = copyTemplateAndFillAuto(
        templateId,
        enrichedData,
        fileName.replace('.' + extension, ''),
        OUTPUT_CONFIG.OUTPUT_FOLDER_ID
      );

      if (format === 'pdf') {
        var pdfFile = exportSpreadsheetAsPdf(spreadsheet, fileName, OUTPUT_CONFIG.OUTPUT_FOLDER_ID);
        downloadUrl = getFileDownloadUrl(pdfFile);
        fileId = pdfFile.getId();
      } else {
        downloadUrl = getExcelDownloadUrl(spreadsheet);
        fileId = spreadsheet.getId();
      }
    }

    // 6. 発注データにパスを記録
    order['注文協議書パス'] = downloadUrl;
    OrderRepository.save(order);

    return {
      success: true,
      url: downloadUrl,
      filename: fileName,
      fileId: fileId
    };
  } catch (e) {
    Logger.log('注文協議書出力エラー: ' + e.message);
    return { success: false, error: e.message };
  }
}

/**
 * 注文書を発行（PDF）
 * テンプレートの種類を自動判別してスプレッドシート/ドキュメント方式を選択
 * @param {string} orderId - 発注ID
 * @returns {Object} 結果オブジェクト {success, url, filename, error}
 */
function issueOrderSheet(orderId) {
  try {
    // 1. テンプレートIDの確認
    var templateId = OUTPUT_CONFIG.TEMPLATE_ORDER_SHEET;
    if (!templateId) {
      return {
        success: false,
        error: '注文書テンプレートファイルIDが設定されていません。OUTPUT_CONFIGを確認してください。'
      };
    }

    // 2. テンプレートの種類を判別
    var templateFile = DriveApp.getFileById(templateId);
    var mimeType = templateFile.getMimeType();

    // Googleドキュメントの場合は専用関数を使用
    if (mimeType === MimeType.GOOGLE_DOCS) {
      return issueOrderSheetFromDocument(orderId);
    }

    // 以下はスプレッドシートテンプレートの場合

    // 3. 発注データを取得
    var order = OrderRepository.findById(orderId);
    if (!order) {
      return { success: false, error: '発注データが見つかりません: ' + orderId };
    }

    // 4. 注文書番号を生成（未設定の場合）
    if (!order['注文書番号']) {
      order['注文書番号'] = generateOrderNumber();
    }

    // 5. 注文書発行日を設定
    order['注文書発行日'] = Utilities.formatDate(new Date(), 'Asia/Tokyo', 'yyyy-MM-dd');

    // 6. 協力会社名を取得（表示用）
    var vendorMap = MasterLookup.getVendorMap();
    var vendorId = order['施主名(注文決定業者名)'];
    order['協力会社名_表示'] = vendorMap[vendorId] || vendorId;

    // 7. ファイル名を生成
    var constructionId = order['関連工事番号'] || orderId;
    var fileName = generateFileName('【注文書】', constructionId, 'pdf');

    // 8. テンプレートをコピーしてデータを埋め込む
    var spreadsheet = copyTemplateAndFill(
      templateId,
      order,
      ORDER_SHEET_MAPPINGS,
      fileName.replace('.pdf', ''),
      OUTPUT_CONFIG.OUTPUT_FOLDER_ID
    );

    // 9. PDFとして出力
    var pdfFile = exportSpreadsheetAsPdf(
      spreadsheet,
      fileName,
      OUTPUT_CONFIG.OUTPUT_FOLDER_ID
    );

    // 10. ダウンロードURLを取得
    var downloadUrl = getFileDownloadUrl(pdfFile);

    // 11. 発注データを更新（注文書番号、発行日、パス）
    order['注文書パス'] = downloadUrl;
    OrderRepository.save(order);

    // 12. 一時スプレッドシートを削除（PDFのみ残す）
    // DriveApp.getFileById(spreadsheet.getId()).setTrashed(true);

    return {
      success: true,
      url: downloadUrl,
      filename: fileName,
      fileId: pdfFile.getId(),
      orderNumber: order['注文書番号']
    };
  } catch (e) {
    Logger.log('注文書発行エラー: ' + e.message);
    return { success: false, error: e.message };
  }
}

/**
 * 注文書番号を生成
 * @returns {string} 注文書番号（例: ORD-2026-0001）
 */
function generateOrderNumber() {
  var year = new Date().getFullYear();
  var prefix = 'ORD-' + year + '-';

  // 既存の注文書番号から最大値を取得
  var orders = OrderRepository.findAll();
  var maxNum = 0;

  orders.forEach(function(order) {
    var orderNum = order['注文書番号'];
    if (orderNum && orderNum.indexOf(prefix) === 0) {
      var num = parseInt(orderNum.replace(prefix, ''), 10);
      if (!isNaN(num) && num > maxNum) {
        maxNum = num;
      }
    }
  });

  // 次の番号を生成（4桁ゼロ埋め）
  var nextNum = maxNum + 1;
  return prefix + ('0000' + nextNum).slice(-4);
}

// ========================================
// マスタ管理 API
// ========================================

/**
 * 顧客一覧を取得
 */
function getCustomers(keyword) {
  try {
    var customers = CustomerRepository.search(keyword);

    return customers.map(function(row) {
      return {
        customerId: String(row['顧客ID'] || ''),
        customerName: String(row['顧客名'] || ''),
        furigana: String(row['ふりがな'] || ''),
        address: String(row['住所'] || ''),
        category: String(row['ふりわけ'] || ''),
        phone: String(row['電話番号'] || '')
      };
    });
  } catch (e) {
    Logger.log('顧客取得エラー: ' + e.message + ' / Stack: ' + e.stack);
    Logger.log('顧客取得エラー: ' + e.message);
    return [];
  }
}

/**
 * 顧客マスタ（軽量版）
 * サジェスト・選択肢用に最小限のフィールドのみ返す。
 * 詳細フィールドが必要な場合は getCustomers() を使う。
 */
function getCustomersLight() {
  try {
    var rows = CustomerRepository.findAll();
    return rows.map(function(r) {
      return {
        customerId: String(r['顧客ID'] || ''),
        customerName: String(r['顧客名'] || ''),
        furigana: String(r['ふりがな'] || '')
      };
    });
  } catch (e) {
    Logger.log('getCustomersLight エラー: ' + e.message);
    return [];
  }
}

/**
 * 顧客を保存
 */
function saveCustomer(data) {
  try {
    var sheetData = {
      '顧客ID': data.customerId,
      '顧客名': data.customerName,
      'ふりがな': data.furigana || '',
      '住所': data.address || '',
      'ふりわけ': data.category || '',
      '電話番号': data.phone || ''
    };

    var result = CustomerRepository.save(sheetData);

    // マスタ更新時はキャッシュをクリア
    MasterLookup.clearCache();

    return { success: true, id: result['顧客ID'] };
  } catch (e) {
    Logger.log('顧客保存エラー: ' + e.message);
    return { success: false, error: e.message };
  }
}

/**
 * 協力会社一覧を取得
 */
function getVendors(keyword) {
  try {
    var vendors = VendorRepository.search(keyword);

    return vendors.map(function(row) {
      return formatVendorData(row);
    });
  } catch (e) {
    Logger.log('協力会社取得エラー: ' + e.message);
    return [];
  }
}

/**
 * 協力会社マスタ（軽量版）
 * サジェスト・選択肢用に最小限のフィールドのみ返す。
 * 詳細フィールドが必要な場合は getVendorById() / getVendors() を使う。
 */
function getVendorsLight() {
  try {
    var rows = VendorRepository.findAll();
    return rows.map(function(r) {
      return {
        vendorId: String(r['仕入先コード'] || ''),
        vendorName: String(r['会社名'] || ''),
        shortName: String(r['略称'] || ''),
        kana: String(r['カナ'] || ''),
        vendorRank: String(r['業者ランク'] || '')
      };
    });
  } catch (e) {
    Logger.log('getVendorsLight エラー: ' + e.message);
    return [];
  }
}

/**
 * 協力会社を1件取得（ID指定）
 * フロントエンドでの業者選択時にマスタ情報を自動セットするために使用
 * @param {string} vendorId - 仕入先コード
 * @returns {Object|null} 協力会社情報（許可ステータス計算済み）
 */
function getVendorById(vendorId) {
  try {
    var row = VendorRepository.findById(vendorId);
    if (!row) return null;
    return formatVendorData(row);
  } catch (e) {
    Logger.log('協力会社取得エラー: ' + e.message);
    return null;
  }
}

/**
 * 協力会社データをフォーマット（共通処理）
 * AppSheetの自動計算を再現：
 * - 許可ステータス: IF(OR(ISBLANK([許可期限]), [許可期限] < TODAY()), "なし", "あり")
 * @param {Object} row - スプレッドシートの行データ
 * @returns {Object} フォーマット済みデータ
 */
function formatVendorData(row) {
  // 許可ステータスの自動計算（AppSheet数式の再現）
  var permitExpiry = row['許可期限'];
  var permitStatus = 'なし';

  if (permitExpiry) {
    var expiryDate = permitExpiry instanceof Date ? permitExpiry : new Date(permitExpiry);
    var today = new Date();
    today.setHours(0, 0, 0, 0);

    if (!isNaN(expiryDate.getTime()) && expiryDate >= today) {
      permitStatus = 'あり';
    }
  }

  return {
    vendorId: row['仕入先コード'],
    vendorName: row['会社名'],
    shortName: row['略称'],
    kana: row['カナ'] || '',
    address: row['住所'],
    phone: row['電話番号'],
    fax: row['ファックス番号'] || '',
    representativeName: row['代表者名'] || '',
    contactName: row['担当者名'],
    closingDay: row['締日'],
    paymentMonth: row['支払日（ヶ月後）'],
    paymentDay: row['支払日'],
    cashRatio: row['支払比率(手形以外)'],
    noteRatio: row['支払比率(手形)'],
    noteSite: row['手形サイト'],
    invoiceNumber: row['事業者登録番号'],
    permitType: row['許可業種'],
    permitExpiry: permitExpiry instanceof Date
      ? Utilities.formatDate(permitExpiry, 'Asia/Tokyo', 'yyyy-MM-dd')
      : permitExpiry,
    permitStatus: permitStatus,
    vendorRank: row['業者ランク'],
    safetyFeeRate: row['安推協会費率']
  };
}

/**
 * 協力会社を保存
 */
function saveVendor(data) {
  try {
    var sheetData = {
      '仕入先コード': data.vendorId || '',
      '会社名': data.vendorName || '',
      '略称': data.shortName || '',
      'カナ': data.kana || '',
      '住所': data.address || '',
      '電話番号': data.phone || '',
      'ファックス番号': data.fax || '',
      '代表者名': data.representativeName || '',
      '担当者名': data.contactName || '',
      '締日': data.closingDay || '',
      '支払日（ヶ月後）': data.paymentMonth || '',
      '支払日': data.paymentDay || '',
      '支払比率(手形以外)': data.cashRatio || '',
      '支払比率(手形)': data.noteRatio || '',
      '手形サイト': data.noteSite || '',
      '事業者登録番号': data.invoiceNumber || '',
      '許可業種': data.permitType || '',
      '許可期限': data.permitExpiry || '',
      '業者ランク': data.vendorRank || '',
      '安推協会費率': data.safetyFeeRate || ''
    };

    var result = VendorRepository.save(sheetData);

    // マスタ更新時はキャッシュをクリア
    MasterLookup.clearCache();

    return { success: true, id: result['仕入先コード'] };
  } catch (e) {
    Logger.log('協力会社保存エラー: ' + e.message);
    return { success: false, error: e.message };
  }
}

/**
 * 社員一覧を取得
 */
function getEmployees(department) {
  try {
    var employees;
    if (department) {
      employees = EmployeeRepository.findByDepartment(department);
    } else {
      employees = EmployeeRepository.findAll();
    }

    return employees.map(function(row) {
      return {
        employeeId: row['社員番号'],
        employeeName: row['氏名'],
        department: row['部門'],
        group: row['所属グループ(主務)'],
        email: row['メールアドレス'],
        phone: row['携帯電話番号(社用)'],
        role: row['承認権限レベル']
      };
    });
  } catch (e) {
    Logger.log('社員取得エラー: ' + e.message);
    return [];
  }
}

/**
 * 社員を保存
 */
function saveEmployee(data) {
  try {
    var sheetData = {
      '社員番号': data.employeeId || data['社員番号'],
      '氏名': data.employeeName || data['氏名'],
      'メールアドレス': data.email || data['メールアドレス'],
      '部門': data.department || data['部門'] || '',
      '承認権限レベル': data.role || data['承認権限レベル'] || ''
    };

    var result = EmployeeRepository.save(sheetData);
    return { success: true, id: result['社員番号'] };
  } catch (e) {
    Logger.log('社員保存エラー: ' + e.message);
    return { success: false, error: e.message };
  }
}

// ========================================
// 請求管理 API
// ========================================

/**
 * 請求一覧を取得
 * @param {Object} filter - フィルター条件（status, projectId, customerId）
 * @returns {Array} 請求一覧
 */
function getInvoices(filter) {
  try {
    var invoices;
    filter = filter || {};

    if (filter.projectId) {
      invoices = InvoiceRepository.findByProjectId(filter.projectId);
    } else if (filter.customerId) {
      invoices = InvoiceRepository.findByCustomerId(filter.customerId);
    } else if (filter.status) {
      invoices = InvoiceRepository.findByStatus(filter.status);
    } else {
      invoices = InvoiceRepository.findAll();
    }

    var customerMap = MasterLookup.getCustomerMap();

    return invoices.map(function(row) {
      var customerId = row['請求先顧客ID'];
      return {
        invoiceId: row['請求ID'],
        projectId: row['案件ID'],
        customerId: customerId,
        customerName: customerMap[customerId] || row['請求先顧客名'] || customerId,
        invoiceType: row['請求区分'],
        invoiceDate: formatDateForApi(row['請求日']),
        amountExcludingTax: Number(row['請求金額(税抜)'] || 0),
        tax: Number(row['消費税'] || 0),
        amountIncludingTax: Number(row['請求金額(税込)'] || 0),
        dueDate: formatDateForApi(row['入金予定日']),
        paymentStatus: row['入金ステータス'] || '未入金',
        paymentDate: formatDateForApi(row['入金日']),
        paymentAmount: Number(row['入金金額'] || 0),
        paymentDifference: Number(row['入金差額'] || 0),
        notes: row['備考'] || ''
      };
    }).reverse();
  } catch (e) {
    Logger.log('請求取得エラー: ' + e.message);
    return [];
  }
}

/**
 * 工事IDに紐づく請求一覧を取得
 * @param {string} projectId - 工事ID
 * @returns {Array} 請求一覧
 */
function getInvoicesByProjectId(projectId) {
  return getInvoices({ projectId: projectId });
}

/**
 * 請求を保存（新規/更新）
 * @param {Object} data - 請求データ
 * @returns {Object} 保存結果
 */
function saveInvoice(data) {
  try {
    // 消費税と税込金額を自動計算
    var amountExcludingTax = Number(data.amountExcludingTax || 0);
    var tax = Math.floor(amountExcludingTax * 0.1);
    var amountIncludingTax = amountExcludingTax + tax;

    var sheetData = {
      '請求ID': data.invoiceId || '',
      '案件ID': data.projectId || '',
      '請求先顧客ID': data.customerId || '',
      '請求先顧客名': data.customerName || '',
      '請求区分': data.invoiceType || '',
      '請求日': data.invoiceDate || '',
      '請求金額(税抜)': amountExcludingTax,
      '消費税': tax,
      '請求金額(税込)': amountIncludingTax,
      '入金予定日': data.dueDate || '',
      '入金ステータス': data.paymentStatus || '未入金',
      '入金日': data.paymentDate || '',
      '入金金額': Number(data.paymentAmount || 0),
      '入金差額': Number(data.paymentDifference || 0),
      '備考': data.notes || ''
    };

    var result = InvoiceRepository.save(sheetData);
    return { success: true, id: result['請求ID'] };
  } catch (e) {
    Logger.log('請求保存エラー: ' + e.message);
    return { success: false, error: e.message };
  }
}

/**
 * 入金消込（入金データを登録）
 * @param {string} invoiceId - 請求ID
 * @param {Object} paymentData - 入金データ
 * @returns {Object} 更新結果
 */
function recordInvoicePayment(invoiceId, paymentData) {
  try {
    var invoice = InvoiceRepository.findById(invoiceId);
    if (!invoice) {
      return { success: false, error: '請求が見つかりません: ' + invoiceId };
    }

    var thisAmount = Number(paymentData.thisAmount || 0);
    var paymentDate = paymentData.paymentDate || formatDateForApi(new Date());

    // 入金履歴を保存
    var historyData = {
      '請求ID': invoiceId,
      '入金日': paymentDate,
      '入金金額': thisAmount,
      '備考': paymentData.notes || ''
    };
    InvoicePaymentHistoryRepository.save(historyData);

    // 履歴から合計入金額を再計算
    var histories = InvoicePaymentHistoryRepository.findByInvoiceId(invoiceId);
    var totalPaid = histories.reduce(function(sum, h) {
      return sum + (Number(h['入金金額']) || 0);
    }, 0);

    var invoiceAmount = Number(invoice['請求金額(税込)'] || 0);
    var difference = invoiceAmount - totalPaid;

    invoice['入金日'] = paymentDate;
    invoice['入金金額'] = totalPaid;
    invoice['入金差額'] = difference;
    invoice['入金ステータス'] = difference <= 0 ? '入金済' : (totalPaid > 0 ? '一部入金' : '未入金');

    var result = InvoiceRepository.save(invoice);
    return { success: true, id: result['請求ID'] };
  } catch (e) {
    Logger.log('入金消込エラー: ' + e.message);
    return { success: false, error: e.message };
  }
}

/**
 * 入金履歴を取得
 * @param {string} invoiceId - 請求ID
 * @returns {Array} 入金履歴一覧
 */
function getInvoicePaymentHistory(invoiceId) {
  try {
    var histories = InvoicePaymentHistoryRepository.findByInvoiceId(invoiceId);
    return histories.map(function(row) {
      return {
        historyId: row['入金履歴ID'] || '',
        invoiceId: row['請求ID'] || '',
        paymentDate: formatDateForApi(row['入金日']),
        paymentAmount: Number(row['入金金額']) || 0,
        notes: row['備考'] || '',
        createdAt: formatDateForApi(row['作成日'])
      };
    });
  } catch (e) {
    Logger.log('入金履歴取得エラー: ' + e.message);
    return [];
  }
}

/**
 * 入金履歴を削除し、請求の入金合計を再計算
 * @param {string} historyId - 入金履歴ID
 * @returns {Object} 削除結果
 */
function deleteInvoicePaymentHistory(historyId) {
  try {
    var history = InvoicePaymentHistoryRepository.findById(historyId);
    if (!history) {
      return { success: false, error: '入金履歴が見つかりません: ' + historyId };
    }
    var invoiceId = history['請求ID'];

    // 履歴を論理削除
    InvoicePaymentHistoryRepository.softDelete(historyId);

    // 残りの履歴から合計入金額を再計算
    var remainingHistories = InvoicePaymentHistoryRepository.findByInvoiceId(invoiceId);
    var totalPaid = remainingHistories.reduce(function(sum, h) {
      return sum + (Number(h['入金金額']) || 0);
    }, 0);

    var invoice = InvoiceRepository.findById(invoiceId);
    if (invoice) {
      var invoiceAmount = Number(invoice['請求金額(税込)'] || 0);
      var difference = invoiceAmount - totalPaid;
      invoice['入金金額'] = totalPaid;
      invoice['入金差額'] = difference;
      invoice['入金ステータス'] = difference <= 0 ? '入金済' : (totalPaid > 0 ? '一部入金' : '未入金');
      if (totalPaid === 0) invoice['入金日'] = '';
      InvoiceRepository.save(invoice);
    }

    return { success: true };
  } catch (e) {
    Logger.log('入金履歴削除エラー: ' + e.message);
    return { success: false, error: e.message };
  }
}

/**
 * 請求ステータスを更新
 * @param {string} invoiceId - 請求ID
 * @param {string} status - 新しいステータス
 * @returns {Object} 更新結果
 */
function updateInvoiceStatus(invoiceId, status) {
  try {
    var invoice = InvoiceRepository.findById(invoiceId);
    if (!invoice) {
      return { success: false, error: '請求が見つかりません: ' + invoiceId };
    }

    invoice['入金ステータス'] = status;
    var result = InvoiceRepository.save(invoice);
    return { success: true, id: result['請求ID'] };
  } catch (e) {
    Logger.log('請求ステータス更新エラー: ' + e.message);
    return { success: false, error: e.message };
  }
}

// ========================================
// 支払管理 API
// ========================================

/**
 * 支払一覧を取得
 * @param {Object} filter - フィルター条件（status, projectId, vendorId, yearMonth）
 * @returns {Array} 支払一覧
 */
function getPayments(filter) {
  try {
    Logger.log('getPayments開始 - filter: ' + JSON.stringify(filter));
    var payments;
    filter = filter || {};

    if (filter.orderId) {
      payments = PaymentRepository.findByOrderId(filter.orderId);
    } else if (filter.projectId) {
      payments = PaymentRepository.findByProjectId(filter.projectId);
    } else if (filter.vendorId) {
      payments = PaymentRepository.findByVendorId(filter.vendorId);
    } else if (filter.yearMonth) {
      payments = PaymentRepository.findByMonth(filter.yearMonth);
    } else if (filter.status) {
      payments = PaymentRepository.findByStatus(filter.status);
    } else {
      Logger.log('PaymentRepository.findAll()を呼び出し');
      payments = PaymentRepository.findAll();
    }

    Logger.log('取得した支払データ件数: ' + (payments ? payments.length : 'null'));

    var vendorMap = MasterLookup.getVendorMap();

    // 名前解決マップを作成
    // 案件ID → 案件名
    var projectMap = {};
    try {
      var allProjects = ProjectRepository.findAll();
      allProjects.forEach(function(p) {
        projectMap[p['案件ID']] = p['案件名'] || '';
      });
    } catch (e) {
      Logger.log('案件マップ作成スキップ: ' + e.message);
    }
    // 工事番号 → 工事名（支払管理の案件ID列には工事番号が入っているケースがある）
    var constructionMap = {};
    try {
      var allConstructions = ConstructionRepository.findAll();
      allConstructions.forEach(function(c) {
        constructionMap[c['工事番号']] = c['工事名'] || '';
      });
    } catch (e) {
      Logger.log('工事マップ作成スキップ: ' + e.message);
    }

    // 支払履歴を一括取得してpaymentIdでグルーピング（N+1回避）
    var allHistory = [];
    try {
      allHistory = PaymentHistoryRepository.findAll();
    } catch (e) {
      Logger.log('支払履歴取得スキップ（シート未作成の可能性）: ' + e.message);
    }
    var historyByPaymentId = {};
    allHistory.forEach(function(h) {
      var pid = String(h['支払ID'] || '').trim();
      if (!historyByPaymentId[pid]) historyByPaymentId[pid] = [];
      historyByPaymentId[pid].push(h);
    });

    return payments.map(function(row) {
      var vendorId = row['協力会社ID'];
      var paymentId = row['支払ID'];
      // 請求月をYYYY-MM形式に変換
      var invoiceMonth = row['請求月'];
      if (invoiceMonth instanceof Date) {
        invoiceMonth = Utilities.formatDate(invoiceMonth, 'Asia/Tokyo', 'yyyy-MM');
      } else if (invoiceMonth) {
        invoiceMonth = String(invoiceMonth);
      }

      var amountIncludingTax = Number(row['請求金額(税込)'] || 0);
      var status = row['支払ステータス'] || '未払';

      // 支払済合計を計算
      var histories = historyByPaymentId[paymentId] || [];
      var totalPaid = 0;
      if (histories.length > 0) {
        histories.forEach(function(h) {
          totalPaid += Number(h['支払金額_合計'] || 0);
        });
      } else if (status === '支払済') {
        // フォールバック: ステータスが支払済だが履歴なし → 全額支払済として表示
        totalPaid = amountIncludingTax;
      }
      var remainingAmount = amountIncludingTax - totalPaid;

      var projectId = row['案件ID'] || '';
      return {
        paymentId: paymentId,
        orderId: row['発注ID'],
        projectId: projectId,
        projectName: projectMap[projectId] || constructionMap[projectId] || projectId,
        vendorId: vendorId,
        vendorName: vendorMap[vendorId] || row['協力会社名'] || vendorId,
        invoiceMonth: invoiceMonth,
        closingDay: row['締日'],
        dueDate: formatDateForApi(row['支払予定日']),
        amountExcludingTax: Number(row['請求金額(税抜)'] || 0),
        tax: Number(row['消費税'] || 0),
        amountIncludingTax: amountIncludingTax,
        progressConfirmed: row['出来高確認'] === 'はい' || row['出来高確認'] === true,
        progressAmount: Number(row['出来高金額'] || 0),
        cashPayment: Number(row['支払方法_現金'] || 0),
        notePayment: Number(row['支払方法_手形'] || 0),
        noteSite: row['手形サイト'] || '',
        paymentStatus: status,
        paymentDate: formatDateForApi(row['支払実行日']),
        notes: row['備考'] || '',
        orderCategory: row['発注区分'] || 'その他',
        safetyFee: Number(row['安全協力会費'] || 0),
        paymentBase: Number(row['支払算定基礎額'] || 0),
        costAmount: Number(row['原価(税抜)'] || 0),
        settlementType: row['決済条件'] || '通常',
        safetyFeeExempt: row['安推協対象外'] === 'はい',
        totalPaid: totalPaid,
        remainingAmount: remainingAmount
      };
    }).reverse();
  } catch (e) {
    Logger.log('支払取得エラー: ' + e.message);
    return [];
  }
}

/**
 * 発注IDに紐づく支払一覧を取得
 * @param {string} orderId - 発注ID
 * @returns {Array} 支払一覧
 */
function getPaymentsByOrderId(orderId) {
  return getPayments({ orderId: orderId });
}

/**
 * 月次の支払一覧を取得
 * @param {string} yearMonth - 年月（YYYY-MM形式）
 * @returns {Array} 支払一覧
 */
function getPaymentsByMonth(yearMonth) {
  return getPayments({ yearMonth: yearMonth });
}

/**
 * 支払を保存（新規/更新）
 * @param {Object} data - 支払データ
 * @returns {Object} 保存結果
 */
function savePayment(data) {
  try {
    // 消費税と税込金額を自動計算
    var amountExcludingTax = Number(data.amountExcludingTax || 0);
    var tax = Math.floor(amountExcludingTax * 0.1);
    var amountIncludingTax = amountExcludingTax + tax;

    var sheetData = {
      '支払ID': data.paymentId || '',
      '発注ID': data.orderId || '',
      '案件ID': data.projectId || '',
      '協力会社ID': data.vendorId || '',
      '協力会社名': data.vendorName || '',
      '請求月': data.invoiceMonth || '',
      '締日': data.closingDay || '',
      '支払予定日': data.dueDate || '',
      '請求金額(税抜)': amountExcludingTax,
      '消費税': tax,
      '請求金額(税込)': amountIncludingTax,
      '出来高確認': data.progressConfirmed ? 'はい' : 'いいえ',
      '出来高金額': Number(data.progressAmount || 0),
      '支払方法_現金': Number(data.cashPayment || 0),
      '支払方法_手形': Number(data.notePayment || 0),
      '手形サイト': data.noteSite || '',
      '支払ステータス': data.paymentStatus || '未払',
      '支払実行日': data.paymentDate || '',
      '備考': data.notes || '',
      '発注区分': data.orderCategory || 'その他',
      '安全協力会費': Number(data.safetyFee || 0),
      '支払算定基礎額': Number(data.paymentBase || 0),
      '原価(税抜)': amountExcludingTax,
      '決済条件': data.settlementType || '通常',
      '安推協対象外': data.safetyFeeExempt ? 'はい' : 'いいえ'
    };

    var result = PaymentRepository.save(sheetData);
    return { success: true, id: result['支払ID'] };
  } catch (e) {
    Logger.log('支払保存エラー: ' + e.message);
    return { success: false, error: e.message };
  }
}

/**
 * 出来高査定（出来高金額を確認・承認）
 * @param {string} paymentId - 支払ID
 * @param {number} approvedAmount - 承認金額
 * @returns {Object} 更新結果
 */
function approvePayment(paymentId, approvedAmount) {
  try {
    var payment = PaymentRepository.findById(paymentId);
    if (!payment) {
      return { success: false, error: '支払が見つかりません: ' + paymentId };
    }

    payment['出来高確認'] = 'はい';
    payment['出来高金額'] = Number(approvedAmount || 0);
    payment['支払ステータス'] = '承認済';

    var result = PaymentRepository.save(payment);
    return { success: true, id: result['支払ID'] };
  } catch (e) {
    Logger.log('出来高査定エラー: ' + e.message);
    return { success: false, error: e.message };
  }
}

/**
 * 支払実行を記録
 * 支払履歴に追加し、支払管理のステータスを自動更新
 * @param {Object} data - 支払実行データ
 * @returns {Object} 結果
 */
function recordPaymentExecution(data) {
  try {
    var paymentId = data.paymentId;
    if (!paymentId) {
      return { success: false, error: '支払IDが指定されていません' };
    }

    var payment = PaymentRepository.findById(paymentId);
    if (!payment) {
      return { success: false, error: '支払が見つかりません: ' + paymentId };
    }

    var cashAmount = Number(data.cashAmount || 0);
    var noteAmount = Number(data.noteAmount || 0);
    var totalAmount = cashAmount + noteAmount;

    if (totalAmount <= 0) {
      return { success: false, error: '支払金額を入力してください' };
    }

    // 残額チェック: 既存履歴の合計を取得
    var existingHistory = PaymentHistoryRepository.findByPaymentId(paymentId);
    var existingTotal = 0;
    existingHistory.forEach(function(h) {
      existingTotal += Number(h['支払金額_合計'] || 0);
    });
    var amountIncludingTax = Number(payment['請求金額(税込)'] || 0);
    var remaining = amountIncludingTax - existingTotal;

    if (totalAmount > remaining + 1) { // 1円の丸め誤差を許容
      return { success: false, error: '支払金額(' + totalAmount + ')が残額(' + remaining + ')を超えています' };
    }

    // 支払履歴レコードを作成
    var historyData = {
      '支払ID': paymentId,
      '発注ID': payment['発注ID'] || '',
      '支払実行日': data.executionDate || '',
      '支払金額_現金': cashAmount,
      '支払金額_手形': noteAmount,
      '支払金額_合計': totalAmount,
      '手形サイト': data.noteSite || '',
      '手形期日': data.noteDueDate || '',
      '備考': data.notes || '',
      '作成日': new Date().toISOString()
    };

    var result = PaymentHistoryRepository.save(historyData);

    // 支払管理のステータスを自動更新
    var newTotal = existingTotal + totalAmount;
    if (newTotal >= amountIncludingTax - 1) { // 1円の丸め誤差を許容
      payment['支払ステータス'] = '支払済';
      payment['支払実行日'] = data.executionDate || '';
    } else {
      payment['支払ステータス'] = '一部支払済';
    }
    PaymentRepository.save(payment);

    return { success: true, id: result['支払履歴ID'], paymentStatus: payment['支払ステータス'] };
  } catch (e) {
    Logger.log('支払実行記録エラー: ' + e.message);
    return { success: false, error: e.message };
  }
}

/**
 * 支払IDの履歴一覧を取得
 * @param {string} paymentId - 支払ID
 * @returns {Array} 支払履歴一覧
 */
function getPaymentHistory(paymentId) {
  try {
    var histories = PaymentHistoryRepository.findByPaymentId(paymentId);
    return histories.map(function(row) {
      return {
        historyId: row['支払履歴ID'],
        paymentId: row['支払ID'],
        orderId: row['発注ID'],
        executionDate: formatDateForApi(row['支払実行日']),
        cashAmount: Number(row['支払金額_現金'] || 0),
        noteAmount: Number(row['支払金額_手形'] || 0),
        totalAmount: Number(row['支払金額_合計'] || 0),
        noteSite: row['手形サイト'] || '',
        noteDueDate: formatDateForApi(row['手形期日']),
        notes: row['備考'] || '',
        createdAt: row['作成日'] || ''
      };
    });
  } catch (e) {
    Logger.log('支払履歴取得エラー: ' + e.message);
    return [];
  }
}

/**
 * 支払履歴を削除（間違い修正用・ソフトデリート）
 * 削除後、支払管理のステータスを再計算
 * @param {string} historyId - 支払履歴ID
 * @returns {Object} 結果
 */
function deletePaymentHistory(historyId) {
  try {
    var history = PaymentHistoryRepository.findById(historyId);
    if (!history) {
      return { success: false, error: '支払履歴が見つかりません: ' + historyId };
    }

    var paymentId = history['支払ID'];

    // ソフトデリート
    history['削除フラグ'] = 'TRUE';
    history['削除日時'] = new Date().toISOString();
    history['削除者'] = Session.getActiveUser().getEmail() || 'system';
    updateDataWithAudit(PaymentHistoryRepository.SHEET_NAME, PaymentHistoryRepository.ID_COLUMN, historyId, history);

    // 支払管理のステータスを再計算
    if (paymentId) {
      var payment = PaymentRepository.findById(paymentId);
      if (payment) {
        var remainingHistory = PaymentHistoryRepository.findByPaymentId(paymentId);
        var totalPaid = 0;
        remainingHistory.forEach(function(h) {
          totalPaid += Number(h['支払金額_合計'] || 0);
        });
        var amountIncludingTax = Number(payment['請求金額(税込)'] || 0);

        if (totalPaid <= 0) {
          // 履歴が全てなくなった → 承認済に戻す（出来高承認済みの場合）
          payment['支払ステータス'] = payment['出来高確認'] === 'はい' ? '承認済' : '未払';
          payment['支払実行日'] = '';
        } else if (totalPaid >= amountIncludingTax - 1) {
          payment['支払ステータス'] = '支払済';
        } else {
          payment['支払ステータス'] = '一部支払済';
        }
        PaymentRepository.save(payment);
      }
    }

    return { success: true };
  } catch (e) {
    Logger.log('支払履歴削除エラー: ' + e.message);
    return { success: false, error: e.message };
  }
}

/**
 * 協力会社マスタの支払条件を適用（でんさい・安推協会費・振込の計算）
 *
 * 計算式:
 *   でんさい額 = 税抜金額 × 50%（万円未満切捨て）
 *   安推協会費 = 税抜金額 × 0.2%（10円未満切捨て）
 *   消費税    = 税抜金額 × 10%（円未満切捨て）
 *   振込金額  = 税込金額 − でんさい額 − 安推協会費（帳尻合わせ）
 *
 * @param {string} vendorId - 仕入先コード
 * @param {number} baseAmount - 基本金額（税抜）
 * @param {string} orderCategory - 発注区分（'材料' or 'その他'）
 * @returns {Object} 支払条件が適用されたデータ
 */
function applyVendorPaymentTerms(vendorId, baseAmount, orderCategory) {
  try {
    var vendor = VendorRepository.findById(vendorId);
    if (!vendor) {
      return { success: false, error: '協力会社が見つかりません: ' + vendorId };
    }

    var amountExcludingTax = Number(baseAmount || 0);

    // 原価 = 税抜き請求金額（安全協力会費・消費税控除前）
    var costAmount = amountExcludingTax;

    // 安推協会費: 税抜金額 × 0.2%（10円未満切捨て）
    // 協力会社マスタの値を優先（%表記）、なければデフォルト0.2%
    var masterRate = parseFloat(vendor['安推協会費率']);
    var safetyFeeRate;
    if (!isNaN(masterRate) && masterRate > 0) {
      safetyFeeRate = masterRate / 100; // %→小数変換: 0.2% → 0.002
    } else {
      safetyFeeRate = 0.002; // デフォルト0.2%
    }
    var safetyFee = Math.floor(amountExcludingTax * safetyFeeRate / 10) * 10; // 10円未満切捨て

    // 消費税・税込金額
    var tax = Math.floor(amountExcludingTax * 0.1);
    var amountIncludingTax = amountExcludingTax + tax;

    // 手形額: 税抜金額 × 比率（万円未満切捨て）
    var noteRatio = Number(vendor['支払比率(手形)'] || 50);
    var notePayment = Math.floor(amountExcludingTax * noteRatio / 100 / 10000) * 10000;

    // 振込金額: 税込金額 - 手形額 - 安推協会費（帳尻合わせ）
    var cashPayment = amountIncludingTax - notePayment - safetyFee;

    return {
      success: true,
      data: {
        vendorId: vendorId,
        vendorName: vendor['会社名'],
        closingDay: vendor['締日'],
        paymentMonth: vendor['支払日（ヶ月後）'],
        paymentDay: vendor['支払日'],
        noteRatio: noteRatio,
        noteSite: vendor['手形サイト'] || '',
        amountExcludingTax: amountExcludingTax,
        tax: tax,
        amountIncludingTax: amountIncludingTax,
        safetyFeeRate: safetyFeeRate,
        safetyFee: safetyFee,
        costAmount: costAmount,
        cashPayment: cashPayment,
        notePayment: notePayment
      }
    };
  } catch (e) {
    Logger.log('支払条件適用エラー: ' + e.message);
    return { success: false, error: e.message };
  }
}

/**
 * 月次・業者別の支払集計を計算
 * 条件②: 決済条件による通常/全額振込のグループ分け
 * 条件③: 月次税抜合計10万円未満は全額振込
 *
 * @param {string} vendorId - 協力会社ID
 * @param {string} yearMonth - 年月（YYYY-MM形式）
 * @returns {Object} 月次集計結果
 */
function calculateMonthlyVendorPayment(vendorId, yearMonth) {
  try {
    // 該当業者・該当月の支払レコードを取得
    var allPayments = getPayments({ yearMonth: yearMonth });
    var vendorPayments = allPayments.filter(function(p) {
      return p.vendorId === vendorId;
    });

    if (vendorPayments.length === 0) {
      return { success: true, data: null };
    }

    // 決済条件でグループ分け
    var normalGroup = [];   // 通常（でんさい対象）
    var fullCashGroup = []; // 全額振込
    vendorPayments.forEach(function(p) {
      if (p.settlementType === '全額振込') {
        fullCashGroup.push(p);
      } else {
        normalGroup.push(p);
      }
    });

    // 通常グループの税抜合計
    var normalTotalExTax = 0;
    var normalTotalInTax = 0;
    var normalSafetyFee = 0;
    normalGroup.forEach(function(p) {
      normalTotalExTax += p.amountExcludingTax;
      normalTotalInTax += p.amountIncludingTax;
      if (!p.safetyFeeExempt) {
        normalSafetyFee += p.safetyFee;
      }
    });

    // 全額振込グループの税込合計
    var fullCashTotalInTax = 0;
    fullCashGroup.forEach(function(p) {
      fullCashTotalInTax += p.amountIncludingTax;
    });

    // 通常グループの手形・現金を各レコードの保存値から合算
    var normalNoteTotal = 0;
    var normalCashTotal = 0;
    normalGroup.forEach(function(p) {
      normalNoteTotal += (p.notePayment || 0);
      normalCashTotal += (p.cashPayment || 0);
    });

    // 条件③: 通常グループの税抜合計が10万円未満 → でんさい0（全額振込扱い）
    var notePayment = 0;
    var isSmallAmount = normalTotalExTax < 100000;
    if (!isSmallAmount && normalTotalExTax > 0) {
      // 各支払レコードに保存された手形額を合算
      notePayment = normalNoteTotal;
    }

    // 振込金額: 通常グループ税込合計 - 手形 - 安推協 + 全額振込グループ税込合計
    var cashPayment = normalTotalInTax - notePayment - normalSafetyFee + fullCashTotalInTax;

    // 全体合計
    var grandTotalExTax = normalTotalExTax + fullCashGroup.reduce(function(sum, p) { return sum + p.amountExcludingTax; }, 0);
    var grandTotalInTax = normalTotalInTax + fullCashTotalInTax;

    return {
      success: true,
      data: {
        vendorId: vendorId,
        vendorName: vendorPayments[0].vendorName,
        yearMonth: yearMonth,
        // 全体
        totalExcludingTax: grandTotalExTax,
        totalIncludingTax: grandTotalInTax,
        // 通常グループ
        normalCount: normalGroup.length,
        normalTotalExTax: normalTotalExTax,
        normalTotalInTax: normalTotalInTax,
        // 全額振込グループ
        fullCashCount: fullCashGroup.length,
        fullCashTotalInTax: fullCashTotalInTax,
        // 計算結果
        notePayment: notePayment,
        safetyFee: normalSafetyFee,
        cashPayment: cashPayment,
        isSmallAmount: isSmallAmount,
        // 明細
        payments: vendorPayments
      }
    };
  } catch (e) {
    Logger.log('月次業者別支払集計エラー: ' + e.message);
    return { success: false, error: e.message };
  }
}

/**
 * 指定月の全業者の月次支払集計を取得
 * @param {string} yearMonth - 年月（YYYY-MM形式）
 * @returns {Object} 業者別集計一覧
 */
function getMonthlyPaymentSummary(yearMonth) {
  try {
    var allPayments = getPayments({ yearMonth: yearMonth });
    if (allPayments.length === 0) {
      return { success: true, data: [] };
    }

    // 業者IDでグルーピング
    var vendorIds = [];
    allPayments.forEach(function(p) {
      if (vendorIds.indexOf(p.vendorId) === -1) {
        vendorIds.push(p.vendorId);
      }
    });

    var summaries = vendorIds.map(function(vid) {
      var result = calculateMonthlyVendorPayment(vid, yearMonth);
      return result.success ? result.data : null;
    }).filter(function(s) { return s !== null; });

    return { success: true, data: summaries };
  } catch (e) {
    Logger.log('月次支払サマリ取得エラー: ' + e.message);
    return { success: false, error: e.message };
  }
}

// ========================================
// 実行予算管理 API
// ========================================

/**
 * 実行予算一覧を取得
 * @param {Object} filter - フィルター条件（projectId, workId）
 * @returns {Array} 予算一覧
 */
function getBudgets(filter) {
  try {
    var budgets;
    filter = filter || {};

    if (filter.projectId) {
      budgets = BudgetRepository.findByProjectId(filter.projectId);
    } else if (filter.workId) {
      budgets = BudgetRepository.findByWorkId(filter.workId);
    } else {
      budgets = BudgetRepository.findAll();
    }

    return budgets.map(function(row) {
      var estimateAmount = Number(row['見積金額(契約時)'] || 0);
      var budgetAmount = Number(row['実行予算額'] || 0);
      var orderedAmount = Number(row['発注済額'] || 0);
      var difference = budgetAmount - orderedAmount;
      var reductionRate = budgetAmount > 0 ? ((difference / budgetAmount) * 100).toFixed(1) : 0;

      return {
        budgetId: row['予算ID'],
        projectId: row['案件ID'],
        workId: row['工事ID'],
        workType: row['工事種別'],
        estimateAmount: estimateAmount,
        budgetAmount: budgetAmount,
        orderedAmount: orderedAmount,
        difference: difference,
        reductionRate: reductionRate,
        notes: row['備考'] || '',
        createdAt: formatDateForApi(row['作成日']),
        updatedAt: formatDateForApi(row['更新日'])
      };
    }).reverse();
  } catch (e) {
    Logger.log('予算取得エラー: ' + e.message);
    return [];
  }
}

/**
 * 工事IDに紐づく実行予算一覧を取得
 * @param {string} projectId - 工事ID
 * @returns {Array} 予算一覧
 */
function getBudgetByProjectId(projectId) {
  return getBudgets({ projectId: projectId });
}

/**
 * 工事全体の予算サマリーを取得
 * @param {string} projectId - 工事ID
 * @returns {Object} サマリー情報
 */
function getBudgetSummary(projectId) {
  try {
    var summary = BudgetRepository.getBudgetSummaryByProjectId(projectId);

    // 案件情報も取得
    var project = ProjectRepository.findById(projectId);

    return {
      success: true,
      projectId: projectId,
      projectName: project ? project['案件名'] : '',
      contractAmount: project ? Number(project['契約金額'] || 0) : 0,
      totalEstimate: summary.totalEstimate,
      totalBudget: summary.totalBudget,
      totalOrdered: summary.totalOrdered,
      totalDifference: summary.totalDifference,
      averageReductionRate: summary.averageReductionRate,
      itemCount: summary.itemCount
    };
  } catch (e) {
    Logger.log('予算サマリー取得エラー: ' + e.message);
    return { success: false, error: e.message };
  }
}

/**
 * 実行予算を保存（新規/更新）
 * @param {Object} data - 予算データ
 * @returns {Object} 保存結果
 */
function saveBudget(data) {
  try {
    var sheetData = {
      '予算ID': data.budgetId || '',
      '案件ID': data.projectId || '',
      '工事ID': data.workId || '',
      '工事種別': data.workType || '',
      '見積金額(契約時)': Number(data.estimateAmount || 0),
      '実行予算額': Number(data.budgetAmount || 0),
      '発注済額': Number(data.orderedAmount || 0),
      '差額': Number(data.budgetAmount || 0) - Number(data.orderedAmount || 0),
      '原価低減率': data.reductionRate || '',
      '備考': data.notes || ''
    };

    var result = BudgetRepository.save(sheetData);
    return { success: true, id: result['予算ID'] };
  } catch (e) {
    Logger.log('予算保存エラー: ' + e.message);
    return { success: false, error: e.message };
  }
}

/**
 * 発注済額を自動更新（発注データから集計）
 * @param {string} projectId - 工事ID
 * @returns {Object} 更新結果
 */
function updateBudgetActual(projectId) {
  try {
    // 工事別の発注合計を集計
    var constructions = ConstructionRepository.findByProjectId(projectId);
    var workOrderTotals = {};

    constructions.forEach(function(construction) {
      var workId = construction['工事番号'];
      var orders = OrderRepository.findByConstructionId(workId);
      var total = 0;

      orders.forEach(function(order) {
        total += Number(order['注文金額'] || 0);
      });

      workOrderTotals[workId] = total;
    });

    // 予算データを更新
    var budgets = BudgetRepository.findByProjectId(projectId);
    var updatedCount = 0;

    budgets.forEach(function(budget) {
      var workId = budget['工事ID'];
      if (workOrderTotals.hasOwnProperty(workId)) {
        budget['発注済額'] = workOrderTotals[workId];
        budget['差額'] = Number(budget['実行予算額'] || 0) - workOrderTotals[workId];
        BudgetRepository.save(budget);
        updatedCount++;
      }
    });

    return { success: true, updatedCount: updatedCount };
  } catch (e) {
    Logger.log('予算実績更新エラー: ' + e.message);
    return { success: false, error: e.message };
  }
}

// ========================================
// リレーションAPI（データ連携）
// ========================================

/**
 * 工事IDに紐づく発注一覧を取得
 * 協力会社名はID→名前変換済み
 * @param {string} projectId 工事ID
 * @returns {Array} 発注一覧
 */
function getOrdersByProjectId(projectId) {
  try {
    // 工事IDから工事番号を取得
    var constructions = ConstructionRepository.findByProjectId(projectId);
    if (constructions.length === 0) {
      return [];
    }

    var vendorMap = MasterLookup.getVendorMap();
    var allOrders = [];

    constructions.forEach(function(construction) {
      var orders = OrderRepository.findByConstructionId(construction['工事番号']);
      orders.forEach(function(row) {
        var vendorId = row['施主名(注文決定業者名)'];
        allOrders.push({
          orderId: row['発注ID'],
          constructionId: row['関連工事番号'],
          vendorId: vendorId,
          vendorName: vendorMap[vendorId] || vendorId,
          orderContent: row['注文内容'],
          orderAmount: Number(row['注文金額'] || 0),
          budgetAmount: Number(row['実行予算金額'] || 0),
          paymentStatus: row['支払ステータス(協力会社への)']
        });
      });
    });

    return allOrders;
  } catch (e) {
    Logger.log('工事別発注取得エラー: ' + e.message);
    return [];
  }
}

/**
 * 協力会社IDに紐づく発注履歴を取得
 * @param {string} vendorId 仕入先コード
 * @returns {Array} 発注一覧
 */
function getOrdersByVendorId(vendorId) {
  try {
    var allOrders = OrderRepository.findAll();

    return allOrders.filter(function(order) {
      return order['施主名(注文決定業者名)'] === vendorId;
    }).map(function(row) {
      return {
        orderId: row['発注ID'],
        constructionId: row['関連工事番号'],
        orderContent: row['注文内容'],
        orderAmount: row['注文金額'] || 0,
        orderDate: row['注文書発行日'],
        paymentStatus: row['支払ステータス(協力会社への)']
      };
    });
  } catch (e) {
    Logger.log('協力会社別発注取得エラー: ' + e.message);
    return [];
  }
}

// ========================================
// Excel/PDF出力 共通ヘルパー関数
// ========================================

/**
 * ファイル名を生成
 * @param {string} prefix - プレフィックス（【契約】等）
 * @param {string} projectName - 工事名
 * @param {string} extension - 拡張子（xlsx/pdf）
 * @returns {string} 生成されたファイル名
 */
function generateFileName(prefix, projectName, extension) {
  var date = Utilities.formatDate(new Date(), 'Asia/Tokyo', 'yyyyMMdd');
  // ファイル名に使えない文字をエスケープ
  var safeName = projectName.replace(/[\/\\:*?"<>|]/g, '_');
  return prefix + safeName + '_' + date + '.' + extension;
}

/**
 * テンプレートをコピーしてデータを埋め込む（セルマッピング方式）
 * @param {string} templateId - テンプレートファイルのID
 * @param {Object} data - 埋め込むデータ
 * @param {Object} mappings - セル位置とデータキーのマッピング {セル位置: データキー}
 * @param {string} fileName - 新しいファイル名
 * @param {string} folderId - 出力先フォルダID
 * @returns {Spreadsheet} コピーされたスプレッドシート
 */
function copyTemplateAndFill(templateId, data, mappings, fileName, folderId) {
  // テンプレートをコピー
  var templateFile = DriveApp.getFileById(templateId);
  var folder = folderId ? DriveApp.getFolderById(folderId) : DriveApp.getRootFolder();
  var copiedFile = templateFile.makeCopy(fileName, folder);

  // スプレッドシートを開く
  var ss = SpreadsheetApp.openById(copiedFile.getId());
  var sheet = ss.getSheets()[0]; // 最初のシートを対象

  // データを埋め込む
  for (var cellAddress in mappings) {
    var dataKey = mappings[cellAddress];
    var value = data[dataKey];
    if (value !== undefined && value !== null) {
      sheet.getRange(cellAddress).setValue(value);
    }
  }

  // 変更を保存
  SpreadsheetApp.flush();

  return ss;
}

/**
 * テンプレートをコピーしてプレースホルダーを置換（AppSheet形式対応）
 * スプレッドシート内の <<[フィールド名]>> を自動置換
 * @param {string} templateId - テンプレートファイルのID
 * @param {Object} data - 埋め込むデータ
 * @param {string} fileName - 新しいファイル名
 * @param {string} folderId - 出力先フォルダID
 * @returns {Spreadsheet} コピーされたスプレッドシート
 */
function copyTemplateAndFillAuto(templateId, data, fileName, folderId) {
  // テンプレートをコピー
  var templateFile = DriveApp.getFileById(templateId);
  var folder = folderId ? DriveApp.getFolderById(folderId) : DriveApp.getRootFolder();
  var copiedFile = templateFile.makeCopy(fileName, folder);

  // スプレッドシートを開く
  var ss = SpreadsheetApp.openById(copiedFile.getId());

  // 全シートに対してプレースホルダー置換を実行
  var sheets = ss.getSheets();
  sheets.forEach(function(sheet) {
    replaceSpreadsheetPlaceholders(sheet, data);
  });

  // 変更を保存
  SpreadsheetApp.flush();

  return ss;
}

/**
 * スプレッドシートのシート内のプレースホルダーを置換
 * AppSheetのループ構文 <<Start: [Related ...]>> ... <<End>> にも対応
 * @param {Sheet} sheet - シート
 * @param {Object} data - 置換データ
 */
function replaceSpreadsheetPlaceholders(sheet, data) {
  var range = sheet.getDataRange();
  var values = range.getValues();
  var modified = false;

  // === Phase 1: <<Start: [Related ...]>> ループ処理 ===
  for (var i = 0; i < values.length; i++) {
    for (var j = 0; j < values[i].length; j++) {
      var cellStr = String(values[i][j] || '');
      var startMatch = cellStr.match(/<<Start:\s*\[Related\s+(.+?)\]>>/);
      if (!startMatch) continue;

      // ループ開始を検出。テンプレート行のフィールドマッピングを解析
      var quotes = data['見積もり一覧'] || [];
      var templateRow = values[i].slice(); // テンプレート行をコピー

      // テンプレート行の各セルから <<[field]>> パターンを抽出
      var fieldMap = []; // [{col, field, fullTemplate}]
      for (var k = 0; k < templateRow.length; k++) {
        var tmpl = String(templateRow[k] || '');
        // <<Start:...>> を含むセル → 見積業者名フィールドとして扱う
        if (tmpl.match(/<<Start:/)) {
          fieldMap.push({col: k, field: '_vendorName'});
        }
        // <<End>> のみ or <<End>> を含むセル → クリア対象
        else if (tmpl.match(/<<End>>/)) {
          // <<End>> 以外にプレースホルダーがあればそれも処理
          var endFieldMatch = tmpl.replace(/<<End>>/, '').match(/<<\[([^\]]+)\]>>/);
          if (endFieldMatch) {
            fieldMap.push({col: k, field: endFieldMatch[1], hasEnd: true});
          } else {
            fieldMap.push({col: k, field: '_end'});
          }
        }
        // 通常の <<[field]>> プレースホルダー
        else if (tmpl.match(/<<\[/)) {
          var fMatch = tmpl.match(/<<\[([^\]]+)\]>>/);
          if (fMatch) {
            fieldMap.push({col: k, field: fMatch[1]});
          }
        }
      }

      // 各見積もりデータで行を埋める（①②③④の固定行数分）
      for (var q = 0; q < quotes.length && (i + q) < values.length; q++) {
        var quote = quotes[q];
        var prefix = '見積もり' + (q + 1) + '.';

        for (var f = 0; f < fieldMap.length; f++) {
          var col = fieldMap[f].col;
          var field = fieldMap[f].field;

          if (field === '_vendorName') {
            // 協力会社名を解決（enrichDataで変換済みのものを使用）
            values[i + q][col] = data[prefix + '会社名'] || data[prefix + '協力会社名'] || quote['見積業者名'] || '';
          } else if (field === '_end') {
            values[i + q][col] = '';
          } else {
            var val = quote[field];
            if (val !== undefined && val !== null) {
              values[i + q][col] = formatValueForDocument(field, val);
            } else {
              values[i + q][col] = '';
            }
          }
          modified = true;
        }
      }

      // 見積もりがない場合でもマーカーをクリア
      if (quotes.length === 0) {
        for (var f = 0; f < fieldMap.length; f++) {
          values[i][fieldMap[f].col] = '';
          modified = true;
        }
      }
    }
  }

  // === Phase 2: 通常の <<[field]>> プレースホルダー置換 ===
  for (var i = 0; i < values.length; i++) {
    for (var j = 0; j < values[i].length; j++) {
      var cellValue = values[i][j];
      if (typeof cellValue === 'string' && cellValue.indexOf('<<[') !== -1) {
        var newValue = replacePlaceholdersInString(cellValue, data);
        if (newValue !== cellValue) {
          values[i][j] = newValue;
          modified = true;
        }
      }
    }
  }

  if (modified) {
    range.setValues(values);
  }
}

/**
 * 文字列内のAppSheet形式プレースホルダーを置換
 * 単純形式: <<[フィールド名]>>
 * リレーション形式: <<[テーブル名].[フィールド名]>>
 * @param {string} str - 対象文字列
 * @param {Object} data - 置換データ
 * @returns {string} 置換後の文字列
 */
function replacePlaceholdersInString(str, data) {
  // AppSheetリレーション形式 <<[テーブル].[カラム]>> に対応
  // 正規表現: <<[ + (]以外の文字列 + (].[+ ]以外の文字列)を0回以上) + ]>>
  return str.replace(/<<\[([^\]]+(?:\]\.\[[^\]]+)*)\]>>/g, function(match, fieldName) {
    var value = data[fieldName];
    if (value === undefined || value === null) {
      // データが見つからない場合、元のプレースホルダーを返す（デバッグ用）
      // 本番では空文字に変更可能
      Logger.log('未マッチのプレースホルダー: ' + fieldName);
      return '';
    }
    return formatValueForDocument(fieldName, value);
  });
}

/**
 * スプレッドシートをPDFとして出力
 * @param {Spreadsheet} spreadsheet - スプレッドシート
 * @param {string} fileName - PDFファイル名
 * @param {string} folderId - 出力先フォルダID
 * @returns {File} 作成されたPDFファイル
 */
function exportSpreadsheetAsPdf(spreadsheet, fileName, folderId) {
  var ssId = spreadsheet.getId();
  var sheetId = spreadsheet.getSheets()[0].getSheetId();

  // PDF出力URL
  var url = 'https://docs.google.com/spreadsheets/d/' + ssId + '/export?' +
    'format=pdf' +
    '&gid=' + sheetId +
    '&portrait=true' +          // 縦向き
    '&size=A4' +                // A4サイズ
    '&fitw=true' +              // 幅に合わせる
    '&gridlines=false' +        // グリッド線非表示
    '&printtitle=false' +       // タイトル非表示
    '&sheetnames=false' +       // シート名非表示
    '&pagenum=CENTER' +         // ページ番号中央
    '&fzr=false';               // 固定行非表示

  // PDFを取得
  var token = ScriptApp.getOAuthToken();
  var response = UrlFetchApp.fetch(url, {
    headers: { Authorization: 'Bearer ' + token }
  });
  var pdfBlob = response.getBlob().setName(fileName);

  // フォルダに保存
  var folder = folderId ? DriveApp.getFolderById(folderId) : DriveApp.getRootFolder();
  var pdfFile = folder.createFile(pdfBlob);

  return pdfFile;
}

/**
 * スプレッドシートのExcelダウンロードURLを取得
 * @param {Spreadsheet} spreadsheet - スプレッドシート
 * @returns {string} ダウンロードURL
 */
function getExcelDownloadUrl(spreadsheet) {
  var ssId = spreadsheet.getId();
  return 'https://docs.google.com/spreadsheets/d/' + ssId + '/export?format=xlsx';
}

/**
 * ファイルのダウンロードURLを取得
 * @param {File} file - Driveファイル
 * @returns {string} ダウンロードURL
 */
function getFileDownloadUrl(file) {
  // 一時的に閲覧可能にしてURLを返す
  // 注意: 本番環境ではアクセス権限の管理を検討
  return 'https://drive.google.com/uc?export=download&id=' + file.getId();
}

// ========================================
// Googleドキュメント対応（注文書用）
// ========================================

/**
 * Googleドキュメントをコピーしてプレースホルダーを置換
 * AppSheet形式 <<[フィールド名]>> に対応
 * マッピング定義不要 - データのキー名をそのままプレースホルダー名として使用
 * @param {string} templateId - テンプレートドキュメントのID
 * @param {Object} data - 埋め込むデータ（キー名がプレースホルダー名に対応）
 * @param {Object} placeholders - （省略可）追加のマッピング {プレースホルダー名: データキー}
 * @param {string} fileName - 新しいファイル名
 * @param {string} folderId - 出力先フォルダID
 * @returns {Document} コピーされたGoogleドキュメント
 */
function copyDocumentAndFill(templateId, data, placeholders, fileName, folderId) {
  // テンプレートをコピー
  var templateFile = DriveApp.getFileById(templateId);
  var folder = folderId ? DriveApp.getFolderById(folderId) : DriveApp.getRootFolder();
  var copiedFile = templateFile.makeCopy(fileName, folder);

  // ドキュメントを開く
  var doc = DocumentApp.openById(copiedFile.getId());
  var body = doc.getBody();

  // データのすべてのキーに対してプレースホルダー置換を実行
  for (var key in data) {
    var value = formatValueForDocument(key, data[key]);
    // AppSheet形式のプレースホルダーを置換（<<[key]>> 形式）
    var escapedKey = key.replace(/[.*+?^${}()|[\]\\]/g, '\\$&');
    body.replaceText('<<\\[' + escapedKey + '\\]>>', value);
  }

  // 追加のマッピングがある場合は適用（リレーション参照など）
  if (placeholders) {
    for (var placeholder in placeholders) {
      var dataKey = placeholders[placeholder];
      var value = formatValueForDocument(dataKey, data[dataKey]);
      var escapedPlaceholder = placeholder.replace(/[.*+?^${}()|[\]\\]/g, '\\$&');
      body.replaceText('<<\\[' + escapedPlaceholder + '\\]>>', value);
    }
  }

  // 変更を保存
  doc.saveAndClose();

  return doc;
}

/**
 * ドキュメント出力用に値をフォーマット
 * @param {string} key - フィールド名
 * @param {*} value - 値
 * @returns {string} フォーマットされた文字列
 */
function formatValueForDocument(key, value) {
  // null/undefinedは空文字
  if (value === null || value === undefined) {
    return '';
  }
  // 数値の場合
  if (typeof value === 'number') {
    // 金額フィールドで0の場合は空文字（未入力扱い）
    if (value === 0 && key.indexOf('金額') !== -1) {
      return '';
    }
    // 金額フィールドはカンマ区切り
    if (key.indexOf('金額') !== -1) {
      return value.toLocaleString('ja-JP');
    }
    // パーセントフィールドは%付与
    if (key.indexOf('%') !== -1 || key.indexOf('パーセント') !== -1) {
      return value + '%';
    }
    return String(value);
  }
  // 日付の場合
  if (value instanceof Date) {
    return Utilities.formatDate(value, 'Asia/Tokyo', 'yyyy年MM月dd日');
  }
  return String(value);
}

/**
 * GoogleドキュメントをPDFとして出力
 * @param {Document} document - Googleドキュメント
 * @param {string} fileName - PDFファイル名
 * @param {string} folderId - 出力先フォルダID
 * @returns {File} 作成されたPDFファイル
 */
function exportDocumentAsPdf(document, fileName, folderId) {
  var docId = document.getId();
  var blob = DriveApp.getFileById(docId).getAs('application/pdf');
  blob.setName(fileName);

  var folder = folderId ? DriveApp.getFolderById(folderId) : DriveApp.getRootFolder();
  return folder.createFile(blob);
}

/**
 * 注文書を発行（Googleドキュメント版・PDF出力）
 * Googleドキュメントテンプレートを使用してプレースホルダー置換方式で出力
 * @param {string} orderId - 発注ID
 * @returns {Object} 結果オブジェクト {success, url, filename, error}
 */
function issueOrderSheetFromDocument(orderId) {
  try {
    // 1. テンプレートIDの確認
    var templateId = OUTPUT_CONFIG.TEMPLATE_ORDER_SHEET;
    if (!templateId) {
      return {
        success: false,
        error: '注文書テンプレートファイルIDが設定されていません。OUTPUT_CONFIGを確認してください。'
      };
    }

    // 2. 発注データを取得
    var order = OrderRepository.findById(orderId);
    if (!order) {
      return { success: false, error: '発注データが見つかりません: ' + orderId };
    }

    // 3. 注文書番号を生成（未設定の場合）
    if (!order['注文書番号']) {
      order['注文書番号'] = generateOrderNumber();
    }

    // 4. 注文書発行日を設定
    var issueDate = new Date();
    order['注文書発行日'] = Utilities.formatDate(issueDate, 'Asia/Tokyo', 'yyyy-MM-dd');

    // 5. 関連データを取得してマージ（AppSheetリレーション参照対応）
    var enrichedData = enrichOrderDataForDocument(order);

    // 6. ファイル名を生成
    var constructionId = order['関連工事番号'] || orderId;
    var projectName = enrichedData['関連工事番号.工事名'] || constructionId;
    var pdfFileName = generateFileName('【注文書】', projectName, 'pdf');
    var docFileName = pdfFileName.replace('.pdf', '');

    // 7. テンプレートをコピーしてデータを埋め込む（マッピング不要）
    var document = copyDocumentAndFill(
      templateId,
      enrichedData,
      null,  // 自動マッピングを使用
      docFileName,
      OUTPUT_CONFIG.OUTPUT_FOLDER_ID
    );

    // 8. PDFとして出力
    var pdfFile = exportDocumentAsPdf(
      document,
      pdfFileName,
      OUTPUT_CONFIG.OUTPUT_FOLDER_ID
    );

    // 9. ダウンロードURLを取得
    var downloadUrl = getFileDownloadUrl(pdfFile);

    // 10. 発注データを更新（注文書番号、発行日、パス）
    order['注文書パス'] = downloadUrl;
    OrderRepository.save(order);

    // 11. 一時ドキュメントを削除（PDFのみ残す場合）
    // DriveApp.getFileById(document.getId()).setTrashed(true);

    return {
      success: true,
      url: downloadUrl,
      filename: pdfFileName,
      fileId: pdfFile.getId(),
      orderNumber: order['注文書番号']
    };
  } catch (e) {
    Logger.log('注文書発行エラー（ドキュメント版）: ' + e.message);
    return { success: false, error: e.message };
  }
}

/**
 * 契約協議データにリレーション参照データを追加
 * @param {Object} contract - 契約協議データ
 * @param {string} department - 部署
 * @returns {Object} リレーションデータを含む拡張データ
 */
function enrichContractDataForDocument(contract, department) {
  var data = {};

  // 日付フィールドのリスト（yyyy/MM/dd形式に変換する）
  var dateFields = [
    '作成日', '契約協議書日付', '工期_開始', '工期_終了',
    '入金時期_着工時', '入金時期_中間時', '入金時期_竣工時',
    '承認日'
  ];

  // 元のデータをコピー（日付フィールドはフォーマット変換）
  for (var key in contract) {
    var value = contract[key];

    // 日付フィールドの場合はフォーマット変換
    if (dateFields.indexOf(key) !== -1 && value) {
      data[key] = formatDateValue(value);
    } else {
      data[key] = value;
    }
  }

  // 顧客情報を取得
  var customerId = contract['発注者'];
  if (customerId) {
    try {
      var customer = CustomerRepository.findById(customerId);
      if (customer) {
        // リレーション参照名（発注者 or 契約者）
        var relationName = '発注者';

        for (var cKey in customer) {
          // シンプル形式: 発注者.顧客名
          data[relationName + '.' + cKey] = customer[cKey];
          // AppSheet形式: 発注者].[顧客名 （<<[発注者].[顧客名]>> に対応）
          data[relationName + '].[' + cKey] = customer[cKey];
        }
        // ショートカット
        data['顧客名'] = customer['顧客名'];
        data['顧客名_表示'] = customer['顧客名'];
        data['顧客住所'] = customer['住所'];
        data['顧客電話番号'] = customer['電話番号'];
        data['顧客ふりがな'] = customer['ふりがな'];
      }
    } catch (e) {
      Logger.log('顧客情報取得エラー: ' + e.message);
    }
  }

  // 直接原価_割合と粗付加金額_割合を計算
  var constructionPrice = Number(contract['工事価格']) || 0;
  var netDirectCost = Number(contract['NET_直接原価']) || 0;
  var netGrossProfit = Number(contract['NET_粗付加金額']) || 0;

  if (constructionPrice > 0) {
    data['直接原価_割合'] = ((netDirectCost / constructionPrice) * 100).toFixed(1) + '%';
    data['粗付加金額_割合'] = ((netGrossProfit / constructionPrice) * 100).toFixed(1) + '%';
  } else {
    data['直接原価_割合'] = '0.0%';
    data['粗付加金額_割合'] = '0.0%';
  }

  // 今日の日付
  var today = new Date();
  data['TODAY()'] = Utilities.formatDate(today, 'Asia/Tokyo', 'yyyy年MM月dd日');
  data['今日'] = Utilities.formatDate(today, 'Asia/Tokyo', 'yyyy年MM月dd日');

  return data;
}

/**
 * 日付値をyyyy/MM/dd形式にフォーマット
 * @param {any} value - 日付値（Date, 文字列、など）
 * @returns {string} フォーマット済み日付文字列
 */
function formatDateValue(value) {
  if (!value) return '';

  var date;

  // Dateオブジェクトの場合
  if (value instanceof Date) {
    date = value;
  }
  // 文字列の場合
  else if (typeof value === 'string') {
    // 既にyyyy/MM/dd形式の場合はそのまま返す
    if (/^\d{4}\/\d{2}\/\d{2}$/.test(value)) {
      return value;
    }
    // yyyy-MM-dd形式の場合
    if (/^\d{4}-\d{2}-\d{2}$/.test(value)) {
      return value.replace(/-/g, '/');
    }
    // その他の文字列はDateとしてパース
    date = new Date(value);
  }
  // 数値（シリアル値）の場合
  else if (typeof value === 'number') {
    // Googleスプレッドシートのシリアル値からDateに変換
    date = new Date((value - 25569) * 86400 * 1000);
  }
  else {
    return String(value);
  }

  // 有効な日付かチェック
  if (isNaN(date.getTime())) {
    return String(value);
  }

  // yyyy/MM/dd形式にフォーマット
  return Utilities.formatDate(date, 'Asia/Tokyo', 'yyyy/MM/dd');
}

/**
 * 発注データにリレーション参照データを追加
 * AppSheetの <<[テーブル名].[フィールド名]>> 形式に対応
 * @param {Object} order - 発注データ
 * @returns {Object} リレーションデータを含む拡張データ
 */
function enrichOrderDataForDocument(order) {
  var data = {};

  // 元のデータをコピー
  for (var key in order) {
    data[key] = order[key];
  }

  // 協力会社情報を取得（施主名(注文決定業者名) → 協力会社マスタ）
  var vendorId = order['施主名(注文決定業者名)'];
  if (vendorId) {
    try {
      // まず仕入先コードで検索、見つからなければ会社名で検索（既存データ互換）
      var vendor = VendorRepository.findById(vendorId);
      if (!vendor) {
        var allVendors = VendorRepository.findAll();
        vendor = allVendors.filter(function(v) {
          return String(v['会社名'] || '').trim() === String(vendorId).trim();
        })[0] || null;
      }
      if (vendor) {
        // リレーション参照形式でデータを追加（ドット形式）
        for (var vKey in vendor) {
          data['施主名(注文決定業者名).' + vKey] = vendor[vKey];
        }
        // AppSheet形式（角括弧+ドット+角括弧）にも対応
        for (var vKey in vendor) {
          data['施主名(注文決定業者名)].[' + vKey] = vendor[vKey];
        }
        // よく使うフィールドは短縮名も追加
        data['協力会社名'] = vendor['会社名'];
        data['協力会社名_表示'] = vendor['会社名'];
        data['注文決定業者名'] = vendor['会社名'];
      } else {
        // マスタに見つからない場合、元の値を業者名として使用
        data['協力会社名'] = vendorId;
        data['協力会社名_表示'] = vendorId;
        data['注文決定業者名'] = vendorId;
      }
    } catch (e) {
      Logger.log('協力会社情報取得エラー: ' + e.message);
    }
  }

  // 工事情報を取得（関連工事番号 → 工事管理）
  var constructionId = order['関連工事番号'];
  if (constructionId) {
    try {
      var construction = ConstructionRepository.findById(constructionId);
      if (construction) {
        // リレーション参照形式でデータを追加（ドット形式）
        for (var cKey in construction) {
          data['関連工事番号.' + cKey] = construction[cKey];
        }
        // AppSheet形式（角括弧+ドット+角括弧）にも対応
        for (var cKey in construction) {
          data['関連工事番号].[' + cKey] = construction[cKey];
        }
      }
    } catch (e) {
      Logger.log('工事情報取得エラー: ' + e.message);
    }
  }

  // 見積もり詳細を取得（発注ID → 見積もり詳細）
  // ※相見積もりの場合、複数業者のデータがある
  var orderId = order['発注ID'];
  if (orderId) {
    try {
      var quotes = QuoteRepository.findByOrderId(orderId);
      Logger.log('[enrichOrderData] 発注ID=' + orderId + ' 見積もり件数=' + (quotes ? quotes.length : 0));
      if (quotes && quotes.length > 0) {
        for (var qi = 0; qi < quotes.length; qi++) {
          Logger.log('[enrichOrderData] 見積もり' + (qi + 1) + ': ID=' + quotes[qi]['見積もり詳細ID'] + ', 業者=' + quotes[qi]['見積業者名']);
        }
        // 全ての見積もりデータを配列として追加（相見積もり比較用）
        data['見積もり一覧'] = quotes;

        // 各見積もりのデータを個別に追加（見積もり1、見積もり2、見積もり3...）
        for (var i = 0; i < quotes.length; i++) {
          var quote = quotes[i];
          var prefix = '見積もり' + (i + 1) + '.';

          for (var qKey in quote) {
            data[prefix + qKey] = quote[qKey];
          }

          // 見積業者名を協力会社名に変換（IDまたは名前で検索）
          var quoteVendorId = quote['見積業者名'];
          if (quoteVendorId) {
            var quoteVendor = VendorRepository.findById(quoteVendorId);
            if (!quoteVendor) {
              var allV = VendorRepository.findAll();
              quoteVendor = allV.filter(function(v) {
                return String(v['会社名'] || '').trim() === String(quoteVendorId).trim();
              })[0] || null;
            }
            if (quoteVendor) {
              data[prefix + '会社名'] = quoteVendor['会社名'];
              data[prefix + '協力会社名'] = quoteVendor['会社名'];
              data[prefix + '見積業者名'] = quoteVendor['会社名'];
            } else {
              data[prefix + '会社名'] = quoteVendorId;
              data[prefix + '協力会社名'] = quoteVendorId;
            }
          }
        }

        // テンプレート用キーを出力（デバッグ）
        var quoteKeys = Object.keys(data).filter(function(k) { return k.indexOf('見積もり') === 0 && k.indexOf('.') !== -1; });
        Logger.log('[enrichOrderData] 見積もりプレースホルダーキー: ' + quoteKeys.join(', '));

        // 選定業者（注文決定業者）の見積もりを特定
        var selectedQuote = null;
        for (var i = 0; i < quotes.length; i++) {
          if (quotes[i]['見積業者名'] === vendorId) {
            selectedQuote = quotes[i];
            break;
          }
        }
        // 見つからなければ最初の見積もりを使用
        if (!selectedQuote && quotes.length > 0) {
          selectedQuote = quotes[0];
        }

        if (selectedQuote) {
          // 見積もり詳細のデータを直接追加（テンプレート用）
          data['TEL'] = selectedQuote['TEL'] || '';
          data['担当者'] = selectedQuote['担当者'] || '';
          data['見積金額'] = selectedQuote['見積金額'] || 0;
          data['当社NET'] = selectedQuote['当社NET'] || 0;
          data['最終協議金額'] = selectedQuote['最終協議金額'] || 0;

          // 見積業者名を協力会社名に変換（IDまたは名前で検索）
          var selectedVendorId = selectedQuote['見積業者名'];
          if (selectedVendorId) {
            var selectedVendor = VendorRepository.findById(selectedVendorId);
            if (!selectedVendor) {
              var allV2 = VendorRepository.findAll();
              selectedVendor = allV2.filter(function(v) {
                return String(v['会社名'] || '').trim() === String(selectedVendorId).trim();
              })[0] || null;
            }
            if (selectedVendor) {
              data['見積業者名'] = selectedVendor['会社名'];
            } else {
              data['見積業者名'] = selectedVendorId;
            }
          }

          // リレーション参照形式でも追加
          for (var qKey in selectedQuote) {
            data['見積もり詳細.' + qKey] = selectedQuote[qKey];
          }
        }
      }
    } catch (e) {
      Logger.log('見積もり詳細取得エラー: ' + e.message);
    }
  }

  // 今日の日付（様々な形式）
  var today = new Date();
  data['TODAY()'] = Utilities.formatDate(today, 'Asia/Tokyo', 'yyyy年MM月dd日');
  data['今日'] = Utilities.formatDate(today, 'Asia/Tokyo', 'yyyy年MM月dd日');

  return data;
}

// ========================================
// マッピング定義（参考用・自動マッピングに移行済み）
// ========================================
//
// 現在はAppSheet形式 <<[フィールド名]>> の自動マッピングを使用しています。
// テンプレート内のプレースホルダーとスプレッドシートのカラム名が一致していれば
// 自動的に置換されます。
//
// リレーション参照（<<[関連工事番号].[工事名]>> など）も対応済み：
// - enrichOrderDataForDocument(): 発注→協力会社、発注→工事のリレーション
// - enrichContractDataForDocument(): 契約協議→顧客のリレーション
//
// 【対応しているプレースホルダー形式】
// - <<[フィールド名]>>           → データのカラム名と直接マッチ
// - <<[テーブル名].[フィールド名]>> → リレーション参照（enrich関数で事前展開）
// - <<[TODAY()]>>                → 今日の日付（yyyy年MM月dd日形式）
//
// ========================================

// ========================================
// 開発・デバッグ用
// ========================================

/**
 * 出力設定の状態を確認
 * テンプレートIDとフォルダIDが正しく設定されているかチェック
 */
function checkOutputConfig() {
  Logger.log('=== 出力設定確認 ===');

  var configs = [
    { key: 'TEMPLATE_CONTRACT_SOLUTION', name: '契約協議書テンプレート' },
    { key: 'TEMPLATE_ORDER_DOCUMENT', name: '注文協議書テンプレート' },
    { key: 'TEMPLATE_ORDER_SHEET', name: '注文書テンプレート' },
    { key: 'OUTPUT_FOLDER_ID', name: '出力フォルダ' }
  ];

  var allConfigured = true;

  configs.forEach(function(config) {
    var id = OUTPUT_CONFIG[config.key];
    var status = '未設定';
    var fileInfo = '';

    if (id) {
      try {
        var file = DriveApp.getFileById(id);
        status = '✓ 設定済み';
        fileInfo = ' (' + file.getName() + ' - ' + file.getMimeType() + ')';
      } catch (e) {
        status = '✗ ID無効（ファイルが見つかりません）';
        allConfigured = false;
      }
    } else {
      allConfigured = false;
    }

    Logger.log(config.name + ': ' + status + fileInfo);
  });

  Logger.log('');
  Logger.log(allConfigured ? '全ての設定が完了しています' : '未設定の項目があります');

  return allConfigured;
}

/**
 * 出力機能のテスト実行
 * 引数なしで呼び出すとダミーデータでテスト
 */
function testExportFunctions() {
  Logger.log('=== 出力機能テスト ===');

  // 設定確認
  if (!checkOutputConfig()) {
    Logger.log('テストを中止します。OUTPUT_CONFIGを設定してください。');
    return;
  }

  // テスト用の契約協議IDと発注IDを取得（存在するものを使用）
  var contracts = ContractRepository.findAll();
  var orders = OrderRepository.findAll();

  if (contracts.length > 0) {
    var testContractId = contracts[0]['協議書ID'];
    Logger.log('契約協議書Excel出力テスト: ' + testContractId);
    var result = exportContractToExcel(testContractId, 'ソリューション事業部');
    Logger.log(result.success ? '成功: ' + result.url : 'エラー: ' + result.error);
  } else {
    Logger.log('契約協議データがありません。スキップします。');
  }

  if (orders.length > 0) {
    var testOrderId = orders[0]['発注ID'];
    Logger.log('注文協議書出力テスト: ' + testOrderId);
    var result = exportOrderDocument(testOrderId, 'excel');
    Logger.log(result.success ? '成功: ' + result.url : 'エラー: ' + result.error);
  } else {
    Logger.log('発注データがありません。スキップします。');
  }

  Logger.log('テスト完了');
}

/**
 * 注文書発行のデバッグテスト
 * GASエディタから直接実行して動作確認
 */
function debugIssueOrderSheet() {
  Logger.log('=== 注文書発行デバッグ ===');

  // 1. 設定確認
  Logger.log('【1】設定確認');
  var templateId = OUTPUT_CONFIG.TEMPLATE_ORDER_SHEET;
  var folderId = OUTPUT_CONFIG.OUTPUT_FOLDER_ID;

  Logger.log('テンプレートID: ' + (templateId || '未設定'));
  Logger.log('出力フォルダID: ' + (folderId || '未設定'));

  if (!templateId) {
    Logger.log('エラー: テンプレートIDが未設定です');
    return;
  }

  // 2. テンプレートファイル確認
  Logger.log('【2】テンプレートファイル確認');
  try {
    var templateFile = DriveApp.getFileById(templateId);
    Logger.log('ファイル名: ' + templateFile.getName());
    Logger.log('MIMEタイプ: ' + templateFile.getMimeType());
    Logger.log('Googleドキュメント?: ' + (templateFile.getMimeType() === MimeType.GOOGLE_DOCS));
  } catch (e) {
    Logger.log('エラー: テンプレートファイルにアクセスできません - ' + e.message);
    return;
  }

  // 3. 発注データ確認
  Logger.log('【3】発注データ確認');
  var orders = OrderRepository.findAll();
  Logger.log('発注データ件数: ' + orders.length);

  if (orders.length === 0) {
    Logger.log('エラー: 発注データがありません');
    return;
  }

  var testOrder = orders[0];
  var testOrderId = testOrder['発注ID'];
  Logger.log('テスト発注ID: ' + testOrderId);
  Logger.log('関連工事番号: ' + testOrder['関連工事番号']);
  Logger.log('協力会社ID: ' + testOrder['施主名(注文決定業者名)']);

  // 4. enrichデータ確認
  Logger.log('【4】enrichデータ確認');
  try {
    var enrichedData = enrichOrderDataForDocument(testOrder);
    Logger.log('enrichedDataキー数: ' + Object.keys(enrichedData).length);
    Logger.log('協力会社名: ' + enrichedData['協力会社名']);
    Logger.log('関連工事番号.工事名: ' + enrichedData['関連工事番号.工事名']);
  } catch (e) {
    Logger.log('エラー: enrichOrderDataForDocument - ' + e.message);
    return;
  }

  // 5. 注文書発行テスト
  Logger.log('【5】注文書発行テスト');
  try {
    var result = issueOrderSheet(testOrderId);
    Logger.log('結果: ' + JSON.stringify(result));
    if (result.success) {
      Logger.log('✓ 成功！');
      Logger.log('URL: ' + result.url);
      Logger.log('ファイル名: ' + result.filename);
    } else {
      Logger.log('✗ 失敗: ' + result.error);
    }
  } catch (e) {
    Logger.log('例外エラー: ' + e.message);
    Logger.log('スタックトレース: ' + e.stack);
  }

  Logger.log('=== デバッグ完了 ===');
}

/**
 * 指定した発注IDで注文書発行テスト
 * @param {string} orderId - テストする発注ID
 */
function testIssueOrderSheetById(orderId) {
  if (!orderId) {
    // 発注IDが指定されていない場合は最初のデータを使用
    var orders = OrderRepository.findAll();
    if (orders.length > 0) {
      orderId = orders[0]['発注ID'];
    } else {
      Logger.log('発注データがありません');
      return;
    }
  }

  Logger.log('発注ID: ' + orderId + ' で注文書発行テスト');
  var result = issueOrderSheet(orderId);
  Logger.log(JSON.stringify(result, null, 2));
  return result;
}

/**
 * 注文協議書出力のデバッグテスト
 */
function debugExportOrderDocument() {
  Logger.log('=== 注文協議書出力デバッグ ===');

  // 1. 設定確認
  var templateId = OUTPUT_CONFIG.TEMPLATE_ORDER_DOCUMENT;
  var folderId = OUTPUT_CONFIG.OUTPUT_FOLDER_ID;
  Logger.log('テンプレートID: ' + templateId);
  Logger.log('出力フォルダID: ' + folderId);

  if (!templateId) {
    Logger.log('エラー: テンプレートIDが未設定');
    return;
  }

  // 2. テンプレートファイル確認
  var templateFile;
  try {
    templateFile = DriveApp.getFileById(templateId);
    Logger.log('ファイル名: ' + templateFile.getName());
    Logger.log('MIMEタイプ: ' + templateFile.getMimeType());
    Logger.log('オーナー: ' + templateFile.getOwner().getEmail());
  } catch (e) {
    Logger.log('テンプレートアクセスエラー: ' + e.message);
    return;
  }

  // 3. 出力フォルダ確認
  try {
    var folder = DriveApp.getFolderById(folderId);
    Logger.log('出力フォルダ名: ' + folder.getName());
  } catch (e) {
    Logger.log('出力フォルダアクセスエラー: ' + e.message);
    return;
  }

  // 4. テンプレートコピーテスト
  Logger.log('--- テンプレートコピーテスト ---');
  try {
    var copiedFile = templateFile.makeCopy('テスト_' + new Date().getTime(), folder);
    Logger.log('コピー成功: ' + copiedFile.getId());
    Logger.log('コピーファイルMIME: ' + copiedFile.getMimeType());

    // スプレッドシートとして開けるか
    var ss = SpreadsheetApp.openById(copiedFile.getId());
    Logger.log('スプレッドシートとして開けました: ' + ss.getName());

    // テストファイル削除
    copiedFile.setTrashed(true);
    Logger.log('テストファイル削除完了');
  } catch (e) {
    Logger.log('コピー/オープンエラー: ' + e.message);
    Logger.log('スタック: ' + e.stack);
    return;
  }

  // 5. 発注データ取得
  var orders = OrderRepository.findAll();
  Logger.log('発注データ件数: ' + orders.length);

  if (orders.length === 0) {
    Logger.log('発注データがありません');
    return;
  }

  var testOrderId = orders[0]['発注ID'];
  Logger.log('テスト発注ID: ' + testOrderId);

  // 6. 出力実行
  try {
    var result = exportOrderDocument(testOrderId, 'excel');
    Logger.log('結果: ' + JSON.stringify(result));
  } catch (e) {
    Logger.log('出力エラー: ' + e.message);
    Logger.log('スタック: ' + e.stack);
  }
}

/**
 * スプレッドシートのシート構成を確認
 */
function checkSheetStructure() {
  var ss = SpreadsheetApp.openById(SPREADSHEET_ID);
  var sheets = ss.getSheets();

  Logger.log('=== シート一覧 ===');
  Logger.log('スプレッドシート名: ' + ss.getName());
  Logger.log('シート数: ' + sheets.length);

  for (var i = 0; i < sheets.length; i++) {
    var sheet = sheets[i];
    var sheetName = sheet.getName();
    var lastRow = sheet.getLastRow();
    var lastCol = sheet.getLastColumn();

    Logger.log('【' + (i + 1) + '】 ' + sheetName);
    Logger.log('  データ行数: ' + lastRow);
    Logger.log('  列数: ' + lastCol);

    if (lastCol > 0) {
      var headers = sheet.getRange(1, 1, 1, lastCol).getValues()[0];
      Logger.log('  カラム: ' + headers.join(', '));
    }
    Logger.log('');
  }

  return 'ログを確認してください';
}

/**
 * テンプレート内のプレースホルダーを抽出
 * AppSheet形式のプレースホルダーを一覧表示
 * @param {string} templateType - テンプレート種別（'order_document', 'order_sheet', 'contract_solution', 'contract_housing'）
 */
function extractPlaceholders(templateType) {
  var templateId;
  var templateName;

  switch (templateType) {
    case 'order_document':
      templateId = OUTPUT_CONFIG.TEMPLATE_ORDER_DOCUMENT;
      templateName = '注文協議書';
      break;
    case 'order_sheet':
      templateId = OUTPUT_CONFIG.TEMPLATE_ORDER_SHEET;
      templateName = '注文書';
      break;
    case 'contract_solution':
      templateId = OUTPUT_CONFIG.TEMPLATE_CONTRACT_SOLUTION;
      templateName = '契約協議書';
      break;
    default:
      Logger.log('無効なテンプレート種別: ' + templateType);
      return;
  }

  Logger.log('=== ' + templateName + ' プレースホルダー抽出 ===');
  Logger.log('テンプレートID: ' + templateId);

  try {
    var file = DriveApp.getFileById(templateId);
    var mimeType = file.getMimeType();
    Logger.log('ファイル名: ' + file.getName());
    Logger.log('MIMEタイプ: ' + mimeType);

    var text = '';

    // Googleドキュメントの場合
    if (mimeType === MimeType.GOOGLE_DOCS) {
      var doc = DocumentApp.openById(templateId);
      text = doc.getBody().getText();
    }
    // Googleスプレッドシートの場合
    else if (mimeType === MimeType.GOOGLE_SHEETS) {
      var ss = SpreadsheetApp.openById(templateId);
      var sheets = ss.getSheets();
      for (var i = 0; i < sheets.length; i++) {
        var sheet = sheets[i];
        var range = sheet.getDataRange();
        var values = range.getValues();
        for (var r = 0; r < values.length; r++) {
          for (var c = 0; c < values[r].length; c++) {
            text += values[r][c] + ' ';
          }
        }
      }
    }

    // AppSheet形式のプレースホルダーを抽出
    var placeholderPattern = /<<\[([^\]]+)\]>>/g;
    var placeholders = [];
    var match;
    while ((match = placeholderPattern.exec(text)) !== null) {
      if (placeholders.indexOf(match[1]) === -1) {
        placeholders.push(match[1]);
      }
    }

    Logger.log('');
    Logger.log('【検出されたプレースホルダー】 ' + placeholders.length + '件');
    placeholders.sort();
    for (var i = 0; i < placeholders.length; i++) {
      Logger.log('  ' + (i + 1) + '. <<[' + placeholders[i] + ']>>');
    }

    return placeholders;
  } catch (e) {
    Logger.log('エラー: ' + e.message);
    return [];
  }
}

/**
 * 注文協議書テンプレートのプレースホルダーと利用可能データを比較
 */
function debugOrderDocumentMapping() {
  Logger.log('=== 注文協議書マッピングデバッグ ===');

  // 1. テンプレートからプレースホルダーを抽出
  Logger.log('【1】テンプレートのプレースホルダー');
  var placeholders = extractPlaceholders('order_document');

  // 2. サンプル発注データを取得してenrich
  Logger.log('');
  Logger.log('【2】利用可能なデータキー');
  var orders = OrderRepository.findAll();
  if (orders.length === 0) {
    Logger.log('発注データがありません');
    return;
  }

  var testOrder = orders[0];
  var enrichedData = enrichOrderDataForDocument(testOrder);
  var dataKeys = Object.keys(enrichedData).sort();

  Logger.log('enrichedDataキー数: ' + dataKeys.length);
  for (var i = 0; i < dataKeys.length; i++) {
    Logger.log('  ' + dataKeys[i] + ': ' + enrichedData[dataKeys[i]]);
  }

  // 3. マッチング確認
  Logger.log('');
  Logger.log('【3】マッチング状況');
  var matched = [];
  var unmatched = [];

  for (var i = 0; i < placeholders.length; i++) {
    var ph = placeholders[i];
    if (enrichedData.hasOwnProperty(ph)) {
      matched.push(ph);
    } else {
      unmatched.push(ph);
    }
  }

  Logger.log('マッチ済み: ' + matched.length + '件');
  matched.forEach(function(ph) {
    Logger.log('  OK: ' + ph + ' = ' + enrichedData[ph]);
  });

  Logger.log('');
  Logger.log('未マッチ（データ不足）: ' + unmatched.length + '件');
  unmatched.forEach(function(ph) {
    Logger.log('  NG: ' + ph);
  });

  return {
    placeholders: placeholders,
    dataKeys: dataKeys,
    matched: matched,
    unmatched: unmatched
  };
}

/**
 * 契約協議書テンプレートのプレースホルダーと利用可能データを比較
 * GASエディタから実行: debugContractDocumentMapping()
 */
function debugContractDocumentMapping() {
  var templateType = 'contract_solution';
  var departmentName = 'ソリューション事業部';

  Logger.log('=== 契約協議書マッピングデバッグ（' + departmentName + '） ===');

  // 1. テンプレートからプレースホルダーを抽出
  Logger.log('【1】テンプレートのプレースホルダー');
  var placeholders = extractPlaceholders(templateType);

  // 2. サンプル契約協議データを取得してenrich
  Logger.log('');
  Logger.log('【2】利用可能なデータキー');
  var contracts = ContractRepository.findAll();
  if (contracts.length === 0) {
    Logger.log('契約協議データがありません');
    return;
  }

  var sampleContract = contracts[0];
  Logger.log('サンプル協議書ID: ' + sampleContract['協議書ID']);

  var enrichedData = enrichContractDataForDocument(sampleContract, departmentName);
  var dataKeys = Object.keys(enrichedData).sort();

  Logger.log('利用可能なキー数: ' + dataKeys.length);

  // 3. マッチング結果
  Logger.log('');
  Logger.log('【3】マッチング結果');
  var matched = [];
  var unmatched = [];

  for (var i = 0; i < placeholders.length; i++) {
    var ph = placeholders[i];
    if (enrichedData.hasOwnProperty(ph)) {
      matched.push(ph);
    } else {
      unmatched.push(ph);
    }
  }

  Logger.log('マッチ済み: ' + matched.length + '件');
  matched.forEach(function(ph) {
    var value = enrichedData[ph];
    var displayValue = (value === null || value === undefined || value === '') ? '(空)' : String(value).substring(0, 30);
    Logger.log('  OK: <<[' + ph + ']>> → ' + displayValue);
  });

  Logger.log('');
  Logger.log('未マッチ（データ不足）: ' + unmatched.length + '件');
  unmatched.forEach(function(ph) {
    Logger.log('  NG: <<[' + ph + ']>> ← データにこのキーがありません');
  });

  // 4. データキーの一覧（テンプレートに追加できる候補）
  Logger.log('');
  Logger.log('【4】テンプレートに追加可能なフィールド（一部抜粋）');
  var commonFields = ['協議書ID', '作成日', '契約協議書日付', '区分', '作成者', '担当者', '工事名', '工事場所',
                      '発注者', '電話番号', '精算カテゴリー', '工事価格', '消費税', '契約金額',
                      'NET_直接原価', 'NET_粗付加金額', '入金時期_着工時', '入金金額_着工時',
                      '入金時期_中間時', '入金金額_中間時', '入金時期_竣工時', '入金金額_竣工時',
                      '工期_開始', '工期_終了', '備考', 'TODAY()', '今日', '顧客名', '顧客名_表示'];

  commonFields.forEach(function(field) {
    if (enrichedData.hasOwnProperty(field)) {
      var value = enrichedData[field];
      var displayValue = (value === null || value === undefined || value === '') ? '(空)' : String(value).substring(0, 30);
      Logger.log('  <<[' + field + ']>> → ' + displayValue);
    }
  });

  return {
    placeholders: placeholders,
    dataKeys: dataKeys,
    matched: matched,
    unmatched: unmatched,
    sampleData: enrichedData
  };
}

/**
 * 契約協議書のプレースホルダー確認（ショートカット）
 */
function debugContractSolution() {
  return debugContractDocumentMapping();
}

// ========================================
// ダミーデータ生成（開発・テスト用）
// ========================================

/**
 * 全テーブルにダミーデータを一括挿入
 * マスタデータ（顧客、協力会社、社員）を参照して整合性のあるデータを生成
 * GASエディタから実行: generateAllDummyData()
 */
function generateAllDummyData() {
  Logger.log('=== ダミーデータ一括生成開始 ===');

  try {
    // 1. 契約協議書を生成
    var contracts = generateDummyContracts(3);
    Logger.log('契約協議書: ' + contracts.length + '件生成');

    // 2. 工事管理を生成（契約協議書に基づく）
    var projects = generateDummyProjects(contracts);
    Logger.log('工事管理: ' + projects.length + '件生成');

    // 3. 工事管理を生成（契約協議書に基づく）
    var constructions = generateDummyConstructions(contracts);
    Logger.log('工事管理: ' + constructions.length + '件生成');

    // 4. 発注業者管理を生成（工事に基づく）
    var orders = generateDummyOrders(constructions);
    Logger.log('発注業者管理: ' + orders.length + '件生成');

    Logger.log('=== ダミーデータ一括生成完了 ===');
    return {
      contracts: contracts.length,
      projects: projects.length,
      constructions: constructions.length,
      orders: orders.length
    };
  } catch (e) {
    Logger.log('エラー: ' + e.message);
    Logger.log('スタック: ' + e.stack);
    throw e;
  }
}

/**
 * 契約協議書のダミーデータを生成
 * @param {number} count - 生成件数
 * @returns {Array} 生成した契約協議書データ
 */
function generateDummyContracts(count) {
  count = count || 3;
  var results = [];

  // マスタデータを取得
  var customers = CustomerRepository.findAll();
  var employees = EmployeeRepository.findAll();

  if (customers.length === 0) {
    Logger.log('顧客マスタにデータがありません');
    return results;
  }
  if (employees.length === 0) {
    Logger.log('社員マスタにデータがありません');
    return results;
  }

  // 工事名のサンプル
  var projectNames = [
    '新築工事',
    '増築工事',
    '改修工事',
    '設備更新工事',
    '空調設備工事',
    '電気設備工事',
    '外壁改修工事'
  ];

  // 区分のサンプル
  var categories = ['新規', '追加', '変更'];

  var today = new Date();

  for (var i = 0; i < count; i++) {
    // ランダムに顧客・社員を選択
    var customer = customers[Math.floor(Math.random() * customers.length)];
    var employee = employees[Math.floor(Math.random() * employees.length)];

    // 工事価格（1000万〜5億円）
    var constructionPrice = Math.floor((Math.random() * 490000000 + 10000000) / 100000) * 100000;
    var consumptionTax = Math.floor(constructionPrice * 0.1);
    var contractAmount = constructionPrice + consumptionTax;

    // 直接原価（工事価格の60〜85%）
    var netDirectCost = Math.floor(constructionPrice * (0.6 + Math.random() * 0.25));
    var netGrossProfit = constructionPrice - netDirectCost;

    // 工期（今日から1〜3ヶ月後開始、6〜18ヶ月間）
    var startDate = new Date(today);
    startDate.setMonth(startDate.getMonth() + 1 + Math.floor(Math.random() * 3));
    var endDate = new Date(startDate);
    endDate.setMonth(endDate.getMonth() + 6 + Math.floor(Math.random() * 12));

    // 入金条件（3回払い）
    var paymentStart = Math.floor(contractAmount * 0.3);
    var paymentMiddle = Math.floor(contractAmount * 0.3);
    var paymentEnd = contractAmount - paymentStart - paymentMiddle;

    // 入金時期
    var paymentDateStart = new Date(startDate);
    var paymentDateMiddle = new Date(startDate);
    paymentDateMiddle.setMonth(paymentDateMiddle.getMonth() + Math.floor((endDate - startDate) / (1000 * 60 * 60 * 24 * 30) / 2));
    var paymentDateEnd = new Date(endDate);

    var contractId = ContractRepository.generateId();

    var data = {
      '協議書ID': contractId,
      '作成日': formatDateForSheet(today),
      '契約協議書日付': formatDateForSheet(today),
      '区分': categories[Math.floor(Math.random() * categories.length)],
      '作成者': employee['社員番号'],
      '担当者': employee['氏名'],
      '工事名': customer['顧客名'] + ' 様 ' + projectNames[Math.floor(Math.random() * projectNames.length)],
      '工事場所': customer['住所'] || '東京都千代田区',
      '発注者': customer['顧客ID'],
      '電話番号': customer['電話番号'] || '',

      '精算カテゴリー': '契約',
      '工事価格': constructionPrice,
      '消費税': consumptionTax,
      '契約金額': contractAmount,
      'NET_直接原価': netDirectCost,
      'NET_粗付加金額': netGrossProfit,

      '入金時期_着工時': formatDateForSheet(paymentDateStart),
      '入金金額_着工時': paymentStart,
      '入金時期_中間時': formatDateForSheet(paymentDateMiddle),
      '入金金額_中間時': paymentMiddle,
      '入金時期_竣工時': formatDateForSheet(paymentDateEnd),
      '入金金額_竣工時': paymentEnd,
      '入金方法_現金': '100',
      '入金方法_手形': '0',
      '入金方法_サイト': '',

      '工期_開始': formatDateForSheet(startDate),
      '工期_終了': formatDateForSheet(endDate),
      '別途発注工事の有無': '無',
      '別途発注工事_詳細': '',
      '式典の有無': Math.random() > 0.7 ? '有' : '無',
      '技術者資格条件の有無': '無',
      '技術者資格条件_詳細': '',
      '資金調達方法': '自己資金',
      '資金調達方法_融資': '',
      '履行保証の有無': '無',
      '履行保証_詳細': '',
      '完成保証人の有無': '無',
      '完成保証人_詳細': '',

      '契約書(写)の有無': '◯',
      '設計図書の有無': '◯',
      '仕様書の有無': '◯',
      '見積書の有無': '◯',
      '打合せ記録の有無': '◯',
      '現場説明書の有無': '✕',
      '現場写真の有無': '◯',
      '図面の有無': '◯',

      '備考': 'ダミーデータ（自動生成）',
      '承認ステータス': '作成中'
    };

    try {
      // insertDataを直接使用（saveはIDがあると更新扱いになるため）
      insertData('契約協議書', data);
      results.push(data);
      Logger.log('契約協議書作成: ' + contractId + ' - ' + data['工事名']);
    } catch (e) {
      Logger.log('契約協議書作成エラー: ' + e.message);
    }
  }

  return results;
}

/**
 * 工事管理のダミーデータを生成（契約協議書に基づく）
 * @param {Array} contracts - 契約協議書データ配列
 * @returns {Array} 生成した工事データ
 */
function generateDummyProjects(contracts) {
  var results = [];

  contracts.forEach(function(contract) {
    var projectId = ProjectRepository.generateId();

    var data = {
      '案件ID': projectId,
      '案件名': contract['工事名'],
      '顧客ID': contract['発注者'],
      '契約金額': contract['契約金額'],
      'ステータス': '施工中',
      '担当者ID': contract['作成者'],
      '部署': 'ソリューション事業部',
      '作成日': contract['作成日'],
      '更新日': contract['作成日'],
      '関連契約協議書ID': contract['協議書ID'],
      '備考': 'ダミーデータ（自動生成）'
    };

    try {
      // insertDataを直接使用
      insertData('案件管理', data);
      results.push(data);
      Logger.log('工事作成: ' + projectId + ' - ' + data['案件名']);
    } catch (e) {
      Logger.log('工事作成エラー: ' + e.message);
    }
  });

  return results;
}

/**
 * 工事管理のダミーデータを生成（契約協議書に基づく）
 * @param {Array} contracts - 契約協議書データ配列
 * @returns {Array} 生成した工事データ
 */
function generateDummyConstructions(contracts) {
  var results = [];

  // 工事種別のサンプル
  var workTypes = ['電気設備工事', '空調設備工事', '衛生設備工事', '建築工事', '土木工事'];

  contracts.forEach(function(contract) {
    var constructionId = ConstructionRepository.generateId();

    var data = {
      '工事番号': constructionId,
      '工事名': contract['工事名'],
      '工事種別': workTypes[Math.floor(Math.random() * workTypes.length)],
      '関連案件ID': '', // 工事と紐づける場合は後で設定
      '関連契約協議書ID': contract['協議書ID'],
      '顧客ID': contract['発注者'],
      '工事場所': contract['工事場所'],
      '工期_開始': contract['工期_開始'],
      '工期_終了': contract['工期_終了'],
      '契約金額': contract['契約金額'],
      '実行予算': contract['NET_直接原価'],
      '進捗ステータス': '着工準備中',
      '担当者': contract['担当者'],
      '部署': 'ソリューション事業部',
      '備考': 'ダミーデータ（自動生成）'
    };

    try {
      // insertDataを直接使用
      insertData('工事管理', data);
      results.push(data);
      Logger.log('工事作成: ' + constructionId + ' - ' + data['工事名']);
    } catch (e) {
      Logger.log('工事作成エラー: ' + e.message);
    }
  });

  return results;
}

/**
 * 発注業者管理のダミーデータを生成（工事に基づく）
 * @param {Array} constructions - 工事データ配列
 * @returns {Array} 生成した発注データ
 */
function generateDummyOrders(constructions) {
  var results = [];

  // 協力会社マスタを取得
  var vendors = VendorRepository.findAll();
  if (vendors.length === 0) {
    Logger.log('協力会社マスタにデータがありません');
    return results;
  }

  // 注文内容のサンプル
  var orderContents = [
    '電気設備工事一式',
    '空調設備工事一式',
    '配管工事一式',
    '内装工事一式',
    '外構工事一式',
    '基礎工事一式'
  ];

  constructions.forEach(function(construction) {
    // 各工事に1〜3件の発注を作成
    var orderCount = 1 + Math.floor(Math.random() * 3);
    var totalBudget = Number(construction['実行予算']) || Number(construction['契約金額']) * 0.7;
    var remainingBudget = totalBudget;

    for (var i = 0; i < orderCount; i++) {
      var vendor = vendors[Math.floor(Math.random() * vendors.length)];
      var orderId = OrderRepository.generateId();

      // 発注金額（残り予算の30〜70%、最後は全額）
      var orderAmount;
      if (i === orderCount - 1) {
        orderAmount = remainingBudget;
      } else {
        orderAmount = Math.floor(remainingBudget * (0.3 + Math.random() * 0.4) / 10000) * 10000;
        remainingBudget -= orderAmount;
      }

      // 実行予算（発注金額の100〜110%）
      var budgetAmount = Math.floor(orderAmount * (1 + Math.random() * 0.1));

      var data = {
        '発注ID': orderId,
        '関連工事番号': construction['工事番号'],
        '施主名(注文決定業者名)': vendor['仕入先コード'],
        '注文内容': orderContents[Math.floor(Math.random() * orderContents.length)],
        '注文金額': orderAmount,
        '実行予算金額': budgetAmount,
        '注文決定業者のランク': vendor['ランク'] || 'B',

        '支払条件(毎月◯日締め切り)': vendor['締日'] || '25',
        '支払条件(翌月◯日支払)': vendor['支払日'] || '25',
        '支払条件(出来高◯%以内)': orderAmount <= 500000 ? '100' : '90',
        '支払条件(現金%)': vendor['支払比率(手形以外)'] || '100',
        '支払条件(手形%)': vendor['支払比率(手形)'] || '0',
        '支払ステータス(協力会社への)': '未発注',
        '安推協特別会費': '',

        '協議内容': '',
        '注文条件': '',
        'かし担保': '引渡後2年間',
        '建設業の許可': vendor['許可ステータス'] || '',
        '備考': 'ダミーデータ（自動生成）',

        '注文書発行日': '',
        '注文書番号': '',
        '注文協議書パス': '',
        '注文書パス': '',
        '返送された注文請書': ''
      };

      try {
        // insertDataを直接使用
        insertData('発注業者管理', data);
        results.push(data);
        Logger.log('発注作成: ' + orderId + ' - ' + (vendor['会社名'] || vendor['仕入先コード']) + ' ￥' + orderAmount.toLocaleString());
      } catch (e) {
        Logger.log('発注作成エラー: ' + e.message);
      }
    }
  });

  return results;
}

/**
 * 日付をスプレッドシート用にフォーマット
 * @param {Date} date - 日付
 * @returns {string} yyyy/MM/dd形式の文字列
 */
function formatDateForSheet(date) {
  if (!date) return '';
  return Utilities.formatDate(date, 'Asia/Tokyo', 'yyyy/MM/dd');
}

/**
 * ダミーデータを全て削除（注意：本番では使用しないこと）
 * 備考に「ダミーデータ」を含むレコードを削除
 */
function deleteDummyData() {
  Logger.log('=== ダミーデータ削除開始 ===');
  Logger.log('【注意】この操作は元に戻せません');

  var tables = [
    { name: '発注業者管理', repo: OrderRepository },
    { name: '工事管理', repo: ConstructionRepository },
    { name: '案件管理', repo: ProjectRepository },
    { name: '契約協議書', repo: ContractRepository }
  ];

  tables.forEach(function(table) {
    var allData = table.repo.findAll();
    var deleteCount = 0;

    allData.forEach(function(row) {
      if (row['備考'] && row['備考'].indexOf('ダミーデータ') !== -1) {
        try {
          var idColumn = table.repo.ID_COLUMN;
          var id = row[idColumn];
          // 削除処理（実際に削除する場合はコメントアウトを解除）
          // deleteData(table.name, idColumn, id);
          deleteCount++;
          Logger.log('削除対象: ' + table.name + ' - ' + id);
        } catch (e) {
          Logger.log('削除エラー: ' + e.message);
        }
      }
    });

    Logger.log(table.name + ': ' + deleteCount + '件が削除対象');
  });

  Logger.log('=== ダミーデータ削除完了（実際の削除はコメントアウト解除が必要） ===');
}

/**
 * 請求管理シートにダミーデータを挿入
 * GASエディタから直接実行してください
 */
function insertDummyInvoices() {
  var dummyData = [
    {
      '請求ID': 'INV-2026-0001',
      '案件ID': 'PJ-2024-0001',
      '請求先顧客ID': 'CUS-00001',
      '請求先顧客名': '△△建設株式会社',
      '請求区分': '着手金',
      '請求日': '2026-01-15',
      '請求金額(税抜)': 10000000,
      '消費税': 1000000,
      '請求金額(税込)': 11000000,
      '入金予定日': '2026-02-28',
      '入金ステータス': '入金済',
      '入金日': '2026-02-25',
      '入金金額': 11000000,
      '入金差額': 0,
      '備考': ''
    },
    {
      '請求ID': 'INV-2026-0002',
      '案件ID': 'PJ-2024-0001',
      '請求先顧客ID': 'CUS-00001',
      '請求先顧客名': '△△建設株式会社',
      '請求区分': '中間請求',
      '請求日': '2026-02-01',
      '請求金額(税抜)': 20000000,
      '消費税': 2000000,
      '請求金額(税込)': 22000000,
      '入金予定日': '2026-03-31',
      '入金ステータス': '未入金',
      '入金日': '',
      '入金金額': 0,
      '入金差額': 22000000,
      '備考': ''
    },
    {
      '請求ID': 'INV-2026-0003',
      '案件ID': 'PJ-2024-0002',
      '請求先顧客ID': 'CUS-00002',
      '請求先顧客名': '□□工業株式会社',
      '請求区分': '最終請求',
      '請求日': '2026-01-20',
      '請求金額(税抜)': 50000000,
      '消費税': 5000000,
      '請求金額(税込)': 55000000,
      '入金予定日': '2026-02-28',
      '入金ステータス': '一部入金',
      '入金日': '2026-02-28',
      '入金金額': 30000000,
      '入金差額': 25000000,
      '備考': '分割入金'
    }
  ];

  dummyData.forEach(function(data) {
    try {
      insertData('請求管理', data);
      Logger.log('挿入成功: ' + data['請求ID']);
    } catch (e) {
      Logger.log('挿入エラー: ' + e.message);
    }
  });

  Logger.log('請求ダミーデータ挿入完了: ' + dummyData.length + '件');
}

/**
 * 支払管理シートにダミーデータを挿入
 * GASエディタから直接実行してください
 */
function insertDummyPayments() {
  var dummyData = [
    {
      '支払ID': 'PAY-2026-0001',
      '発注ID': 'ORD-2024-0001',
      '案件ID': 'PJ-2024-0001',
      '協力会社ID': 'V001',
      '協力会社名': '○○電気工業',
      '請求月': '2026-01',
      '締日': '25',
      '支払予定日': '2026-02-10',
      '請求金額(税抜)': 5000000,
      '消費税': 500000,
      '請求金額(税込)': 5500000,
      '出来高確認': 'はい',
      '出来高金額': 5500000,
      '支払方法_現金': 5500000,
      '支払方法_手形': 0,
      '手形サイト': '',
      '支払ステータス': '支払済',
      '支払実行日': '2026-02-10',
      '備考': ''
    },
    {
      '支払ID': 'PAY-2026-0002',
      '発注ID': 'ORD-2024-0002',
      '案件ID': 'PJ-2024-0001',
      '協力会社ID': 'V002',
      '協力会社名': '△△空調設備',
      '請求月': '2026-02',
      '締日': '末',
      '支払予定日': '2026-03-25',
      '請求金額(税抜)': 8000000,
      '消費税': 800000,
      '請求金額(税込)': 8800000,
      '出来高確認': 'いいえ',
      '出来高金額': 0,
      '支払方法_現金': 0,
      '支払方法_手形': 0,
      '手形サイト': '',
      '支払ステータス': '未払',
      '支払実行日': '',
      '備考': '出来高確認待ち'
    },
    {
      '支払ID': 'PAY-2026-0003',
      '発注ID': 'ORD-2024-0003',
      '案件ID': 'PJ-2024-0001',
      '協力会社ID': 'V003',
      '協力会社名': '□□衛生工業',
      '請求月': '2026-02',
      '締日': '20',
      '支払予定日': '2026-03-31',
      '請求金額(税抜)': 3000000,
      '消費税': 300000,
      '請求金額(税込)': 3300000,
      '出来高確認': 'はい',
      '出来高金額': 3300000,
      '支払方法_現金': 1650000,
      '支払方法_手形': 1650000,
      '手形サイト': '90日',
      '支払ステータス': '承認済',
      '支払実行日': '',
      '備考': ''
    }
  ];

  dummyData.forEach(function(data) {
    try {
      insertData('支払管理', data);
      Logger.log('挿入成功: ' + data['支払ID']);
    } catch (e) {
      Logger.log('挿入エラー: ' + e.message);
    }
  });

  Logger.log('支払ダミーデータ挿入完了: ' + dummyData.length + '件');
}

/**
 * 実行予算シートにダミーデータを挿入
 * GASエディタから直接実行してください
 */
function insertDummyBudgets() {
  var dummyData = [
    {
      '予算ID': 'BDG-2026-0001',
      '案件ID': 'PJ-2024-0001',
      '工事ID': 'C-00001',
      '工事種別': '電気設備工事',
      '見積金額(契約時)': 15000000,
      '実行予算額': 13500000,
      '発注済額': 12000000,
      '差額': 1500000,
      '原価低減率': '11.1',
      '備考': '',
      '作成日': new Date(),
      '更新日': new Date()
    },
    {
      '予算ID': 'BDG-2026-0002',
      '案件ID': 'PJ-2024-0001',
      '工事ID': 'C-00001',
      '工事種別': '空調設備工事',
      '見積金額(契約時)': 20000000,
      '実行予算額': 18000000,
      '発注済額': 16000000,
      '差額': 2000000,
      '原価低減率': '11.1',
      '備考': '',
      '作成日': new Date(),
      '更新日': new Date()
    },
    {
      '予算ID': 'BDG-2026-0003',
      '案件ID': 'PJ-2024-0001',
      '工事ID': 'C-00002',
      '工事種別': '衛生設備工事',
      '見積金額(契約時)': 8000000,
      '実行予算額': 7200000,
      '発注済額': 5000000,
      '差額': 2200000,
      '原価低減率': '30.6',
      '備考': '原価削減順調',
      '作成日': new Date(),
      '更新日': new Date()
    },
    {
      '予算ID': 'BDG-2026-0004',
      '案件ID': 'PJ-2024-0002',
      '工事ID': 'C-00003',
      '工事種別': '電気設備工事',
      '見積金額(契約時)': 50000000,
      '実行予算額': 45000000,
      '発注済額': 35000000,
      '差額': 10000000,
      '原価低減率': '22.2',
      '備考': '',
      '作成日': new Date(),
      '更新日': new Date()
    }
  ];

  dummyData.forEach(function(data) {
    try {
      insertData('実行予算', data);
      Logger.log('挿入成功: ' + data['予算ID']);
    } catch (e) {
      Logger.log('挿入エラー: ' + e.message);
    }
  });

  Logger.log('予算ダミーデータ挿入完了: ' + dummyData.length + '件');
}

/**
 * 全ダミーデータを一括挿入
 * GASエディタから直接実行してください
 */
function insertAllDummyData() {
  Logger.log('=== ダミーデータ一括挿入開始 ===');

  insertDummyInvoices();
  insertDummyPayments();
  insertDummyBudgets();

  Logger.log('=== ダミーデータ一括挿入完了 ===');
}

// ========================================
// デバッグ・テスト用関数
// ========================================

/**
 * スプレッドシートのシート名一覧を確認
 */
function debugListSheets() {
  var ss = SpreadsheetApp.openById(SPREADSHEET_ID);
  var sheets = ss.getSheets();
  Logger.log('=== スプレッドシートのシート一覧 ===');
  sheets.forEach(function(sheet, index) {
    Logger.log((index + 1) + ': "' + sheet.getName() + '"');
  });
}

/**
 * 支払管理シートのデータを直接確認
 */
function debugPaymentSheet() {
  Logger.log('=== 支払管理シートデバッグ ===');
  try {
    var ss = SpreadsheetApp.openById(SPREADSHEET_ID);
    var sheet = ss.getSheetByName('支払管理');

    if (!sheet) {
      Logger.log('エラー: 「支払管理」シートが見つかりません');
      Logger.log('利用可能なシート:');
      ss.getSheets().forEach(function(s) {
        Logger.log('  - "' + s.getName() + '"');
      });
      return;
    }

    Logger.log('シート発見: 行数=' + sheet.getLastRow() + ', 列数=' + sheet.getLastColumn());

    // ヘッダー確認
    if (sheet.getLastColumn() > 0) {
      var headers = sheet.getRange(1, 1, 1, sheet.getLastColumn()).getValues()[0];
      Logger.log('ヘッダー: ' + JSON.stringify(headers));
    }

    // データ確認
    var data = PaymentRepository.findAll();
    Logger.log('PaymentRepository.findAll() 結果: ' + (data ? data.length : 'null') + '件');
    if (data && data.length > 0) {
      Logger.log('最初のレコード: ' + JSON.stringify(data[0]));
    }

  } catch (e) {
    Logger.log('エラー: ' + e.message);
    Logger.log('スタック: ' + e.stack);
  }
}

/**
 * getPayments APIのテスト
 */
function testGetPayments() {
  Logger.log('=== getPayments APIテスト ===');
  var result = getPayments({});
  Logger.log('結果タイプ: ' + typeof result);
  Logger.log('結果: ' + JSON.stringify(result));
  Logger.log('件数: ' + (result ? result.length : 'null'));
}

/**
 * 見積もりデータ取得のテスト
 * GASエディタから直接実行してください
 */
function testGetQuotes() {
  Logger.log('=== 見積もりデータ取得テスト ===');

  // まず見積もり詳細シートの全データを確認
  var allQuotes = QuoteRepository.findAll();
  Logger.log('見積もり詳細シート全データ: ' + allQuotes.length + '件');
  if (allQuotes.length > 0) {
    Logger.log('最初のレコード: ' + JSON.stringify(allQuotes[0]));
    Logger.log('発注IDの型: ' + typeof allQuotes[0]['発注ID']);
  }

  // 特定の発注IDで検索テスト
  var testOrderId = 'ORD-2026-0001';
  Logger.log('テスト発注ID: ' + testOrderId);

  var quotes = getQuotesByOrderId(testOrderId);
  Logger.log('getQuotesByOrderId結果: ' + quotes.length + '件');
  Logger.log('結果: ' + JSON.stringify(quotes));
}

// ========================================
// 追加ダミーデータ挿入関数
// ========================================

/**
 * 工事管理シートにダミーデータを挿入
 * GASエディタから直接実行してください
 */
function insertDummyProjects() {
  Logger.log('=== 工事管理ダミーデータ挿入開始 ===');

  var dummyData = [
    {
      '案件ID': 'PJ-2026-0001',
      '案件名': '○○ビル新築電気設備工事',
      '顧客ID': 'CUS-00001',
      '顧客名': '△△建設株式会社',
      '振り分けカテゴリー': 'solution',
      '案件ステータス': '施工中',
      '担当者名': '山田 太郎',
      '部署': 'ソリューション事業部',
      '契約金額': 150000000,
      '着工日': '2026-01-15',
      '竣工予定日': '2026-08-31',
      '契約日': '2025-12-20',
      '関連契約協議書ID': 'CT-2025-0001',
      '備考': '大型工事'
    },
    {
      '案件ID': 'PJ-2026-0002',
      '案件名': '□□工場空調設備更新工事',
      '顧客ID': 'CUS-00002',
      '顧客名': '□□工業株式会社',
      '振り分けカテゴリー': 'solution',
      '案件ステータス': '契約済',
      '担当者名': '佐藤 次郎',
      '部署': 'ソリューション事業部',
      '契約金額': 80000000,
      '着工日': '2026-03-01',
      '竣工予定日': '2026-06-30',
      '契約日': '2026-01-15',
      '関連契約協議書ID': 'CT-2026-0001',
      '備考': ''
    },
    {
      '案件ID': 'PJ-2026-0003',
      '案件名': '△△マンション給排水設備改修',
      '顧客ID': 'CUS-00003',
      '顧客名': '△△不動産管理',
      '振り分けカテゴリー': 'ソリューション',
      '案件ステータス': '契約協議中',
      '担当者名': '鈴木 三郎',
      '部署': 'ソリューション事業部',
      '契約金額': 25000000,
      '着工日': '',
      '竣工予定日': '',
      '契約日': '',
      '関連契約協議書ID': '',
      '備考': '見積り提出済み'
    },
    {
      '案件ID': 'PJ-2026-0004',
      '案件名': '◇◇センター消防設備点検',
      '顧客ID': 'CUS-00004',
      '顧客名': '◇◇商事株式会社',
      '振り分けカテゴリー': 'solution',
      '案件ステータス': '完工',
      '担当者名': '田中 四郎',
      '部署': 'ソリューション事業部',
      '契約金額': 5000000,
      '着工日': '2025-11-01',
      '竣工予定日': '2025-12-15',
      '契約日': '2025-10-20',
      '関連契約協議書ID': 'CT-2025-0002',
      '備考': '定期点検'
    },
    {
      '案件ID': 'PJ-2026-0005',
      '案件名': '××オフィスLED照明更新',
      '顧客ID': 'CUS-00005',
      '顧客名': '××株式会社',
      '振り分けカテゴリー': 'solution',
      '案件ステータス': '下書き',
      '担当者名': '高橋 五郎',
      '部署': 'ソリューション事業部',
      '契約金額': 0,
      '着工日': '',
      '竣工予定日': '',
      '契約日': '',
      '関連契約協議書ID': '',
      '備考': '概算見積り作成中'
    }
  ];

  dummyData.forEach(function(data) {
    try {
      insertData('案件管理', data);
      Logger.log('挿入成功: ' + data['案件ID']);
    } catch (e) {
      Logger.log('挿入エラー: ' + e.message);
    }
  });

  Logger.log('工事管理ダミーデータ挿入完了: ' + dummyData.length + '件');
}

/**
 * 工事管理シートにダミーデータを挿入
 * GASエディタから直接実行してください
 */
function insertDummyConstructions() {
  Logger.log('=== 工事管理ダミーデータ挿入開始 ===');

  var dummyData = [
    {
      '工事番号': 'C-00001',
      '工事名': '○○ビル新築 電気設備工事',
      '関連案件ID': 'PJ-2026-0001',
      '関連契約協議書ID': 'CT-2025-0001',
      '工事場所': '東京都港区○○1-2-3',
      '工事担当社員': '山田 太郎',
      '営業担当社員': '佐藤 次郎',
      '工種・工事範囲': '電気設備一式',
      '工期始': '2026-01-15',
      '工期終': '2026-08-31',
      '進捗ステータス': '施工中',
      '部署': 'ソリューション事業部',
      '記録日': new Date(),
      '備考': ''
    },
    {
      '工事番号': 'C-00002',
      '工事名': '○○ビル新築 空調設備工事',
      '関連案件ID': 'PJ-2026-0001',
      '関連契約協議書ID': 'CT-2025-0001',
      '工事場所': '東京都港区○○1-2-3',
      '工事担当社員': '田中 四郎',
      '営業担当社員': '佐藤 次郎',
      '工種・工事範囲': '空調設備一式',
      '工期始': '2026-02-01',
      '工期終': '2026-08-15',
      '進捗ステータス': '施工中',
      '部署': 'ソリューション事業部',
      '記録日': new Date(),
      '備考': ''
    },
    {
      '工事番号': 'C-00003',
      '工事名': '□□工場 空調設備更新工事',
      '関連案件ID': 'PJ-2026-0002',
      '関連契約協議書ID': 'CT-2026-0001',
      '工事場所': '神奈川県横浜市□□区4-5-6',
      '工事担当社員': '佐藤 次郎',
      '営業担当社員': '山田 太郎',
      '工種・工事範囲': '空調設備更新',
      '工期始': '2026-03-01',
      '工期終': '2026-06-30',
      '進捗ステータス': '未着工',
      '部署': 'ソリューション事業部',
      '記録日': new Date(),
      '備考': ''
    },
    {
      '工事番号': 'C-00004',
      '工事名': '△△マンション 給排水設備改修',
      '関連案件ID': 'PJ-2026-0003',
      '関連契約協議書ID': '',
      '工事場所': '東京都世田谷区△△7-8-9',
      '工事担当社員': '鈴木 三郎',
      '営業担当社員': '鈴木 三郎',
      '工種・工事範囲': '給排水設備改修',
      '工期始': '',
      '工期終': '',
      '進捗ステータス': '見積中',
      '部署': 'ソリューション事業部',
      '記録日': new Date(),
      '備考': '見積り中'
    },
    {
      '工事番号': 'C-00005',
      '工事名': '◇◇センター 消防設備点検',
      '関連案件ID': 'PJ-2026-0004',
      '関連契約協議書ID': 'CT-2025-0002',
      '工事場所': '東京都新宿区◇◇10-11-12',
      '工事担当社員': '田中 四郎',
      '営業担当社員': '高橋 五郎',
      '工種・工事範囲': '消防設備定期点検',
      '工期始': '2025-11-01',
      '工期終': '2025-12-15',
      '進捗ステータス': '完了',
      '部署': 'ソリューション事業部',
      '記録日': new Date(),
      '備考': '完了済み'
    }
  ];

  dummyData.forEach(function(data) {
    try {
      insertData('工事管理', data);
      Logger.log('挿入成功: ' + data['工事番号']);
    } catch (e) {
      Logger.log('挿入エラー: ' + e.message);
    }
  });

  Logger.log('工事管理ダミーデータ挿入完了: ' + dummyData.length + '件');
}

/**
 * 契約協議書シートにダミーデータを挿入
 * GASエディタから直接実行してください
 */
function insertDummyContracts() {
  Logger.log('=== 契約協議書ダミーデータ挿入開始 ===');

  var dummyData = [
    {
      '協議書ID': 'CT-2025-0001',
      '作成日': '2025-12-01',
      '契約協議書日付': '2025-12-15',
      '区分': '新規',
      '作成者': '山田 太郎',
      '担当者': '佐藤 次郎',
      '工事名': '○○ビル新築電気設備工事',
      '工事場所': '東京都港区○○1-2-3',
      '発注者': '△△建設株式会社',
      '電話番号': '03-1234-5678',
      '契約金額': 150000000,
      '工期始': '2026-01-15',
      '工期終': '2026-08-31',
      '承認ステータス': '承認済',
      '承認日': '2025-12-18',
      '備考': ''
    },
    {
      '協議書ID': 'CT-2025-0002',
      '作成日': '2025-10-15',
      '契約協議書日付': '2025-10-20',
      '区分': '継続',
      '作成者': '田中 四郎',
      '担当者': '高橋 五郎',
      '工事名': '◇◇センター消防設備点検',
      '工事場所': '東京都新宿区◇◇10-11-12',
      '発注者': '◇◇商事株式会社',
      '電話番号': '03-9876-5432',
      '契約金額': 5000000,
      '工期始': '2025-11-01',
      '工期終': '2025-12-15',
      '承認ステータス': '承認済',
      '承認日': '2025-10-22',
      '備考': '年次点検'
    },
    {
      '協議書ID': 'CT-2026-0001',
      '作成日': '2026-01-10',
      '契約協議書日付': '2026-01-15',
      '区分': '新規',
      '作成者': '佐藤 次郎',
      '担当者': '山田 太郎',
      '工事名': '□□工場空調設備更新工事',
      '工事場所': '神奈川県横浜市□□区4-5-6',
      '発注者': '□□工業株式会社',
      '電話番号': '045-1234-5678',
      '契約金額': 80000000,
      '工期始': '2026-03-01',
      '工期終': '2026-06-30',
      '承認ステータス': '承認済',
      '承認日': '2026-01-18',
      '備考': ''
    },
    {
      '協議書ID': 'CT-2026-0002',
      '作成日': '2026-02-01',
      '契約協議書日付': '',
      '区分': '新規',
      '作成者': '鈴木 三郎',
      '担当者': '鈴木 三郎',
      '工事名': '△△マンション給排水設備改修',
      '工事場所': '東京都世田谷区△△7-8-9',
      '発注者': '△△不動産管理',
      '電話番号': '03-5555-6666',
      '契約金額': 25000000,
      '工期始': '',
      '工期終': '',
      '承認ステータス': '協議中',
      '承認日': '',
      '備考': '価格交渉中'
    },
    {
      '協議書ID': 'CT-2026-0003',
      '作成日': '2026-02-03',
      '契約協議書日付': '',
      '区分': '新規',
      '作成者': '高橋 五郎',
      '担当者': '高橋 五郎',
      '工事名': '××オフィスLED照明更新',
      '工事場所': '東京都中央区××13-14-15',
      '発注者': '××株式会社',
      '電話番号': '03-7777-8888',
      '契約金額': 0,
      '工期始': '',
      '工期終': '',
      '承認ステータス': '作成中',
      '承認日': '',
      '備考': '概算見積り段階'
    }
  ];

  dummyData.forEach(function(data) {
    try {
      insertData('契約協議書', data);
      Logger.log('挿入成功: ' + data['協議書ID']);
    } catch (e) {
      Logger.log('挿入エラー: ' + e.message);
    }
  });

  Logger.log('契約協議書ダミーデータ挿入完了: ' + dummyData.length + '件');
}

/**
 * 発注業者管理シートにダミーデータを挿入
 * GASエディタから直接実行してください
 */
function insertDummyOrders() {
  Logger.log('=== 発注業者管理ダミーデータ挿入開始 ===');

  var dummyData = [
    {
      '発注ID': 'ORD-2026-0001',
      '関連工事番号': 'C-00001',
      '施主名(注文決定業者名)': '○○電気工業株式会社',
      '注文内容': '電気配線工事一式',
      '注文金額': 12000000,
      '実行予算金額': 13500000,
      '支払ステータス(協力会社への)': '支払済',
      '注文決定業者のランク': 'A',
      '発注日': '2026-01-20',
      '納期': '2026-04-30',
      '備考': ''
    },
    {
      '発注ID': 'ORD-2026-0002',
      '関連工事番号': 'C-00001',
      '施主名(注文決定業者名)': '△△設備工業',
      '注文内容': '分電盤製作・設置',
      '注文金額': 3500000,
      '実行予算金額': 4000000,
      '支払ステータス(協力会社への)': '未払',
      '注文決定業者のランク': 'B',
      '発注日': '2026-02-01',
      '納期': '2026-05-15',
      '備考': ''
    },
    {
      '発注ID': 'ORD-2026-0003',
      '関連工事番号': 'C-00002',
      '施主名(注文決定業者名)': '□□空調設備',
      '注文内容': '空調機器設置工事',
      '注文金額': 16000000,
      '実行予算金額': 18000000,
      '支払ステータス(協力会社への)': '一部払',
      '注文決定業者のランク': 'A',
      '発注日': '2026-02-15',
      '納期': '2026-07-31',
      '備考': '分割払い'
    },
    {
      '発注ID': 'ORD-2026-0004',
      '関連工事番号': 'C-00002',
      '施主名(注文決定業者名)': '◇◇ダクト工業',
      '注文内容': 'ダクト製作・設置',
      '注文金額': 5000000,
      '実行予算金額': 5500000,
      '支払ステータス(協力会社への)': '未払',
      '注文決定業者のランク': 'B',
      '発注日': '2026-02-20',
      '納期': '2026-06-30',
      '備考': ''
    },
    {
      '発注ID': 'ORD-2026-0005',
      '関連工事番号': 'C-00003',
      '施主名(注文決定業者名)': '××機械設備',
      '注文内容': '空調設備更新工事',
      '注文金額': 35000000,
      '実行予算金額': 40000000,
      '支払ステータス(協力会社への)': '未払',
      '注文決定業者のランク': 'A',
      '発注日': '2026-02-25',
      '納期': '2026-06-15',
      '備考': '着工前'
    },
    {
      '発注ID': 'ORD-2026-0006',
      '関連工事番号': 'C-00005',
      '施主名(注文決定業者名)': '○○防災設備',
      '注文内容': '消防設備点検',
      '注文金額': 2500000,
      '実行予算金額': 3000000,
      '支払ステータス(協力会社への)': '支払済',
      '注文決定業者のランク': 'A',
      '発注日': '2025-11-05',
      '納期': '2025-12-10',
      '備考': '完了'
    }
  ];

  dummyData.forEach(function(data) {
    try {
      insertData('発注業者管理', data);
      Logger.log('挿入成功: ' + data['発注ID']);
    } catch (e) {
      Logger.log('挿入エラー: ' + e.message);
    }
  });

  Logger.log('発注業者管理ダミーデータ挿入完了: ' + dummyData.length + '件');
}

/**
 * 見積もり詳細シートにダミーデータを挿入
 * GASエディタから直接実行してください
 */
function insertDummyQuotes() {
  Logger.log('=== 見積もり詳細ダミーデータ挿入開始 ===');

  var dummyData = [
    // ORD-2026-0001の見積もり（3社相見積）
    {
      '見積もり詳細ID': 'Q-000001',
      '発注ID': 'ORD-2026-0001',
      '工事番号': 'C-00001',
      '見積業者名': '○○電気工業株式会社',
      'TEL': '03-1111-2222',
      '担当者': '電気 太郎',
      '見積金額': 12500000,
      '当社NET': 12000000,
      '最終協議金額': 12000000,
      '選定結果': '採用',
      '備考': '最安値'
    },
    {
      '見積もり詳細ID': 'Q-000002',
      '発注ID': 'ORD-2026-0001',
      '工事番号': 'C-00001',
      '見積業者名': 'AA電設',
      'TEL': '03-2222-3333',
      '担当者': '設備 二郎',
      '見積金額': 13200000,
      '当社NET': 12800000,
      '最終協議金額': 0,
      '選定結果': '不採用',
      '備考': ''
    },
    {
      '見積もり詳細ID': 'Q-000003',
      '発注ID': 'ORD-2026-0001',
      '工事番号': 'C-00001',
      '見積業者名': 'BB電機工業',
      'TEL': '03-3333-4444',
      '担当者': '工事 三郎',
      '見積金額': 14000000,
      '当社NET': 13500000,
      '最終協議金額': 0,
      '選定結果': '不採用',
      '備考': '納期が合わず'
    },
    // ORD-2026-0003の見積もり（2社相見積）
    {
      '見積もり詳細ID': 'Q-000004',
      '発注ID': 'ORD-2026-0003',
      '工事番号': 'C-00002',
      '見積業者名': '□□空調設備',
      'TEL': '045-1111-2222',
      '担当者': '空調 四郎',
      '見積金額': 17000000,
      '当社NET': 16000000,
      '最終協議金額': 16000000,
      '選定結果': '採用',
      '備考': '実績あり'
    },
    {
      '見積もり詳細ID': 'Q-000005',
      '発注ID': 'ORD-2026-0003',
      '工事番号': 'C-00002',
      '見積業者名': 'CC冷熱工業',
      'TEL': '045-2222-3333',
      '担当者': '冷熱 五郎',
      '見積金額': 18500000,
      '当社NET': 17500000,
      '最終協議金額': 0,
      '選定結果': '不採用',
      '備考': ''
    },
    // ORD-2026-0005の見積もり（見積中）
    {
      '見積もり詳細ID': 'Q-000006',
      '発注ID': 'ORD-2026-0005',
      '工事番号': 'C-00003',
      '見積業者名': '××機械設備',
      'TEL': '044-1111-2222',
      '担当者': '機械 六郎',
      '見積金額': 38000000,
      '当社NET': 35000000,
      '最終協議金額': 35000000,
      '選定結果': '採用',
      '備考': ''
    },
    {
      '見積もり詳細ID': 'Q-000007',
      '発注ID': 'ORD-2026-0005',
      '工事番号': 'C-00003',
      '見積業者名': 'DD設備工業',
      'TEL': '044-2222-3333',
      '担当者': '設備 七郎',
      '見積金額': 42000000,
      '当社NET': 40000000,
      '最終協議金額': 0,
      '選定結果': '不採用',
      '備考': ''
    }
  ];

  dummyData.forEach(function(data) {
    try {
      insertData('見積もり詳細', data);
      Logger.log('挿入成功: ' + data['見積もり詳細ID']);
    } catch (e) {
      Logger.log('挿入エラー: ' + e.message);
    }
  });

  Logger.log('見積もり詳細ダミーデータ挿入完了: ' + dummyData.length + '件');
}

/**
 * 基幹データ（工事・追加工事・契約・発注・見積）を一括挿入
 * GASエディタから直接実行してください
 */
function insertAllCoreData() {
  Logger.log('=== 基幹データ一括挿入開始 ===');

  insertDummyContracts();   // 契約協議書
  insertDummyProjects();    // 工事管理
  insertDummyConstructions(); // 工事管理
  insertDummyOrders();      // 発注業者管理
  insertDummyQuotes();      // 見積もり詳細

  Logger.log('=== 基幹データ一括挿入完了 ===');
}

/**
 * 全シートにダミーデータを一括挿入（基幹 + 請求/支払/予算）
 * GASエディタから直接実行してください
 */
function insertAllData() {
  Logger.log('=== 全ダミーデータ一括挿入開始 ===');

  insertAllCoreData();      // 基幹データ
  insertDummyInvoices();    // 請求管理
  insertDummyPayments();    // 支払管理
  insertDummyBudgets();     // 実行予算

  Logger.log('=== 全ダミーデータ一括挿入完了 ===');
}

// ========================================
// スキーマ管理（列追加）
// ========================================

/**
 * 契約協議書シートに「部署」列を追加
 * GASエディタから一度だけ実行してください
 * 「担当者」列の隣に「部署」列を挿入します
 */
function addDepartmentColumnToContract() {
  var ss = SpreadsheetApp.openById(SPREADSHEET_ID);
  var sheet = ss.getSheetByName('契約協議書');

  if (!sheet) {
    Logger.log('エラー: 契約協議書シートが見つかりません');
    return;
  }

  var headers = sheet.getRange(1, 1, 1, sheet.getLastColumn()).getValues()[0];

  // 「部署」列が既に存在するか確認
  if (headers.indexOf('部署') !== -1) {
    Logger.log('「部署」列は既に存在します（列番号: ' + (headers.indexOf('部署') + 1) + '）');
    return;
  }

  // 「担当者」列の位置を取得
  var managerColIndex = headers.indexOf('担当者');
  if (managerColIndex === -1) {
    Logger.log('エラー: 「担当者」列が見つかりません');
    return;
  }

  // 「担当者」の次（右隣）に列を挿入
  var insertPosition = managerColIndex + 2; // 1-indexed
  sheet.insertColumnAfter(managerColIndex + 1);

  // ヘッダーに「部署」を設定
  sheet.getRange(1, insertPosition).setValue('部署');

  Logger.log('「部署」列を追加しました（列番号: ' + insertPosition + '、「担当者」列の右隣）');
  Logger.log('既存データの部署は空欄です。担当者を再選択すると自動入力されます。');
}

// ========================================
// 監査ログ API
// ========================================

/**
 * 監査ログ一覧を取得
 * @param {Object} filter - フィルター条件
 * @param {string} filter.startDate - 開始日（YYYY-MM-DD）
 * @param {string} filter.endDate - 終了日（YYYY-MM-DD）
 * @param {string} filter.userEmail - ユーザーEmail
 * @param {string} filter.operationType - 操作種別（CREATE/UPDATE/DELETE/EXPORT/IMPORT）
 * @param {string} filter.targetSheet - 対象シート名
 * @param {string} filter.targetId - 対象ID
 * @param {number} filter.limit - 取得件数上限（デフォルト100）
 * @returns {Array} 監査ログ一覧
 */
function getAuditLogs(filter) {
  try {
    filter = filter || {};
    filter.limit = filter.limit || 100;

    var logs = AuditLogRepository.find(filter);

    return logs.map(function(row) {
      return {
        logId: row['ログID'] || '',
        timestamp: formatDateForApi(row['タイムスタンプ']) || '',
        userEmail: row['ユーザーEmail'] || '',
        userName: row['ユーザー名'] || '',
        department: row['部門'] || '',
        operationType: row['操作種別'] || '',
        targetSheet: row['対象シート'] || '',
        targetId: row['対象ID'] || '',
        beforeJson: row['変更前JSON'] || '',
        afterJson: row['変更後JSON'] || '',
        changedFields: row['変更フィールド'] || ''
      };
    });
  } catch (e) {
    Logger.log('監査ログ取得エラー: ' + e.message);
    return [];
  }
}

/**
 * 特定IDの監査ログ履歴を取得
 * @param {string} targetSheet - 対象シート名
 * @param {string} targetId - 対象ID
 * @returns {Array} 該当する監査ログ一覧
 */
function getAuditLogsByTargetId(targetSheet, targetId) {
  return getAuditLogs({
    targetSheet: targetSheet,
    targetId: targetId,
    limit: 50
  });
}

/**
 * 監査ログの操作種別一覧を取得
 * @returns {Array} 操作種別一覧
 */
function getAuditOperationTypes() {
  return ['CREATE', 'UPDATE', 'DELETE', 'EXPORT', 'IMPORT'];
}

/**
 * 監査ログの対象シート一覧を取得
 * @returns {Array} 対象シート一覧
 */
function getAuditTargetSheets() {
  return [
    '案件管理',
    '契約協議書',
    '工事管理',
    '発注業者管理',
    '請求管理',
    '支払管理',
    '実行予算',
    '顧客マスタ',
    '協力会社マスタ',
    '社員マスタ'
  ];
}

// ========================================
// アーカイブ（削除済みデータ）API
// ========================================

/**
 * 削除済みデータを取得
 * @param {string} sheetName - シート名
 * @returns {Array} 削除済みデータ一覧
 */
function getArchivedData(sheetName) {
  try {
    var data = getDeletedData(sheetName);

    return data.map(function(row) {
      // 日付型を文字列に変換
      var result = {};
      Object.keys(row).forEach(function(key) {
        var value = row[key];
        if (value instanceof Date) {
          result[key] = Utilities.formatDate(value, Session.getScriptTimeZone(), 'yyyy-MM-dd HH:mm:ss');
        } else {
          result[key] = value;
        }
      });
      return result;
    });
  } catch (e) {
    Logger.log('削除済みデータ取得エラー: ' + e.message);
    return [];
  }
}

/**
 * 全シートの削除済みデータをサマリーで取得
 * @returns {Object} シート名ごとの削除件数
 */
function getArchivedDataSummary() {
  try {
    var sheets = [
      { name: '案件管理', label: '工事', idColumn: '案件ID' },
      { name: '契約協議書', label: '契約協議書', idColumn: '協議書ID' },
      { name: '工事管理', label: '工事', idColumn: '工事番号' },
      { name: '発注業者管理', label: '発注', idColumn: '発注ID' },
      { name: '請求管理', label: '請求', idColumn: '請求ID' },
      { name: '支払管理', label: '支払', idColumn: '支払ID' },
      { name: '実行予算', label: '実行予算', idColumn: '予算ID' }
    ];

    var summary = sheets.map(function(sheet) {
      var deletedData = getDeletedData(sheet.name);
      return {
        sheetName: sheet.name,
        label: sheet.label,
        idColumn: sheet.idColumn,
        count: deletedData.length
      };
    });

    return {
      success: true,
      summary: summary,
      totalCount: summary.reduce(function(sum, s) { return sum + s.count; }, 0)
    };
  } catch (e) {
    Logger.log('アーカイブサマリー取得エラー: ' + e.message);
    return { success: false, error: e.message };
  }
}

/**
 * データを論理削除
 * @param {string} sheetName - シート名
 * @param {string} idColumn - IDカラム名
 * @param {string} idValue - 削除対象のID値
 * @returns {Object} 結果
 */
function softDelete(sheetName, idColumn, idValue) {
  try {
    softDeleteData(sheetName, idColumn, idValue);
    return { success: true, id: idValue };
  } catch (e) {
    Logger.log('論理削除エラー: ' + e.message);
    return { success: false, error: e.message };
  }
}

/**
 * データを復元
 * @param {string} sheetName - シート名
 * @param {string} idColumn - IDカラム名
 * @param {string} idValue - 復元対象のID値
 * @returns {Object} 結果
 */
function restoreArchivedData(sheetName, idColumn, idValue) {
  try {
    restoreData(sheetName, idColumn, idValue);
    return { success: true, id: idValue };
  } catch (e) {
    Logger.log('データ復元エラー: ' + e.message);
    return { success: false, error: e.message };
  }
}

/**
 * データを完全削除（物理削除）
 * @param {string} sheetName - シート名
 * @param {string} idColumn - IDカラム名
 * @param {string} idValue - 削除対象のID値
 * @returns {Object} 結果
 */
function permanentDelete(sheetName, idColumn, idValue) {
  try {
    hardDeleteData(sheetName, idColumn, idValue);
    return { success: true, id: idValue };
  } catch (e) {
    Logger.log('完全削除エラー: ' + e.message);
    return { success: false, error: e.message };
  }
}

// ========================================
// 添付ファイル管理 API
// ========================================

/**
 * 添付ファイル保存先フォルダIDを設定
 */
var ATTACHMENT_FOLDER_ID = OUTPUT_CONFIG.OUTPUT_FOLDER_ID;

/**
 * エンティティに紐づく添付ファイル一覧を取得
 * @param {string} entityType - エンティティ種別（project/construction/contract/order）
 * @param {string} entityId - エンティティID
 * @param {string} category - カテゴリ（オプション）
 * @returns {Array} 添付ファイル一覧
 */
function getAttachmentsByEntity(entityType, entityId, category) {
  try {
    var attachments;
    if (category) {
      attachments = AttachmentRepository.findByEntityAndCategory(entityType, entityId, category);
    } else {
      attachments = AttachmentRepository.findByEntity(entityType, entityId);
    }

    var result = attachments.map(function(row) {
      return {
        fileId: row['ファイルID'] || '',
        driveFileId: row['GoogleDriveファイルID'] || '',
        fileName: row['ファイル名'] || '',
        mimeType: row['MIMEタイプ'] || '',
        fileSize: row['ファイルサイズ'] || 0,
        entityType: row['関連エンティティ種別'] || '',
        entityId: row['関連エンティティID'] || '',
        category: row['カテゴリ'] || '',
        uploadedAt: row['アップロード日時'] ? String(row['アップロード日時']) : '',
        uploadedBy: row['アップロード者'] || ''
      };
    });

    // GAS の google.script.run は空配列 [] を null にシリアライズするため、
    // JSON文字列として返し、フロントエンドでパースする
    return JSON.stringify(result);
  } catch (e) {
    Logger.log('添付ファイル取得エラー: ' + e.message);
    return JSON.stringify([]);
  }
}

/**
 * 添付ファイルカテゴリ一覧を取得
 * @returns {Array} カテゴリ一覧
 */
function getAttachmentCategories() {
  return ['契約書', '図面', '写真', '見積書', '報告書', 'その他'];
}

/**
 * ファイルをアップロード
 * @param {Object} params - アップロードパラメータ
 * @param {string} params.entityType - エンティティ種別
 * @param {string} params.entityId - エンティティID
 * @param {string} params.category - カテゴリ
 * @param {string} params.fileName - ファイル名
 * @param {string} params.mimeType - MIMEタイプ
 * @param {string} params.base64Data - Base64エンコードされたファイルデータ
 * @returns {Object} 結果
 */
function uploadAttachment(params) {
  try {
    var user = getAuditUserInfo();

    // Base64デコードしてBlobを作成
    var decoded = Utilities.base64Decode(params.base64Data);
    var blob = Utilities.newBlob(decoded, params.mimeType, params.fileName);

    // Google Driveにアップロード
    var folder = DriveApp.getFolderById(ATTACHMENT_FOLDER_ID);
    var file = folder.createFile(blob);
    var driveFileId = file.getId();

    // メタデータを保存
    var attachmentData = {
      'GoogleDriveファイルID': driveFileId,
      'ファイル名': params.fileName,
      'MIMEタイプ': params.mimeType,
      'ファイルサイズ': decoded.length,
      '関連エンティティ種別': params.entityType,
      '関連エンティティID': params.entityId,
      'カテゴリ': params.category || 'その他',
      'アップロード者': user.email
    };

    var saved = AttachmentRepository.save(attachmentData);

    // 監査ログ
    AuditLogRepository.log({
      userEmail: user.email,
      userName: user.name,
      department: user.department,
      operationType: 'CREATE',
      targetSheet: '添付ファイル管理',
      targetId: saved['ファイルID'],
      beforeJson: '',
      afterJson: JSON.stringify(attachmentData),
      changedFields: Object.keys(attachmentData).join(',')
    });

    return {
      success: true,
      fileId: saved['ファイルID'],
      driveFileId: driveFileId,
      fileName: params.fileName,
      url: file.getUrl()
    };
  } catch (e) {
    Logger.log('ファイルアップロードエラー: ' + e.message);
    return { success: false, error: e.message };
  }
}

/**
 * 添付ファイルのダウンロードURLを取得
 * @param {string} fileId - ファイルID
 * @returns {Object} ダウンロード情報
 */
function getAttachmentDownloadUrl(fileId) {
  try {
    var attachment = AttachmentRepository.findById(fileId);
    if (!attachment) {
      return { success: false, error: 'ファイルが見つかりません' };
    }

    var driveFileId = attachment['GoogleDriveファイルID'];
    var file = DriveApp.getFileById(driveFileId);

    return {
      success: true,
      fileId: fileId,
      driveFileId: driveFileId,
      fileName: attachment['ファイル名'],
      url: file.getUrl(),
      downloadUrl: 'https://drive.google.com/uc?export=download&id=' + driveFileId
    };
  } catch (e) {
    Logger.log('ダウンロードURL取得エラー: ' + e.message);
    return { success: false, error: e.message };
  }
}

/**
 * 添付ファイルを削除
 * @param {string} fileId - ファイルID
 * @param {boolean} deleteDriveFile - Driveからも削除するか（デフォルトfalse）
 * @returns {Object} 結果
 */
function deleteAttachment(fileId, deleteDriveFile) {
  try {
    var attachment = AttachmentRepository.findById(fileId);
    if (!attachment) {
      return { success: false, error: 'ファイルが見つかりません' };
    }

    // Driveからも削除する場合
    if (deleteDriveFile) {
      try {
        var driveFile = DriveApp.getFileById(attachment['GoogleDriveファイルID']);
        driveFile.setTrashed(true);
      } catch (e) {
        Logger.log('Driveファイル削除エラー（続行）: ' + e.message);
      }
    }

    // 論理削除
    AttachmentRepository.softDelete(fileId);

    return { success: true, fileId: fileId };
  } catch (e) {
    Logger.log('添付ファイル削除エラー: ' + e.message);
    return { success: false, error: e.message };
  }
}

// ========================================
// データインポート API
// ========================================

/**
 * インポート対象シート設定
 */
var IMPORT_SHEET_CONFIGS = {
  '案件管理': {
    idColumn: '案件ID',
    requiredColumns: ['案件名'],
    validateRow: function(row) {
      if (!row['案件名']) return { valid: false, error: '工事名は必須です' };
      return { valid: true };
    }
  },
  '契約協議書': {
    idColumn: '協議書ID',
    requiredColumns: ['工事名'],
    validateRow: function(row) {
      if (!row['工事名']) return { valid: false, error: '工事名は必須です' };
      return { valid: true };
    }
  },
  '工事管理': {
    idColumn: '工事番号',
    requiredColumns: ['工事名'],
    validateRow: function(row) {
      if (!row['工事名']) return { valid: false, error: '工事名は必須です' };
      return { valid: true };
    }
  },
  '発注業者管理': {
    idColumn: '発注ID',
    requiredColumns: ['発注先業者名'],
    validateRow: function(row) {
      if (!row['発注先業者名']) return { valid: false, error: '発注先業者名は必須です' };
      return { valid: true };
    }
  },
  '顧客マスタ': {
    idColumn: '顧客ID',
    requiredColumns: ['顧客名'],
    validateRow: function(row) {
      if (!row['顧客名']) return { valid: false, error: '顧客名は必須です' };
      return { valid: true };
    }
  },
  '協力会社マスタ': {
    idColumn: '仕入先コード',
    requiredColumns: ['会社名'],
    validateRow: function(row) {
      if (!row['会社名']) return { valid: false, error: '会社名は必須です' };
      return { valid: true };
    }
  }
};

/**
 * インポート可能なシート一覧を取得
 * @returns {Array} シート一覧
 */
function getImportableSheets() {
  try {
    var sheets = Object.keys(IMPORT_SHEET_CONFIGS).map(function(name) {
      var config = IMPORT_SHEET_CONFIGS[name];
      return {
        name: name,
        idColumn: config.idColumn,
        requiredColumns: config.requiredColumns
      };
    });
    return sheets;
  } catch (e) {
    Logger.log('インポート可能シート一覧取得エラー: ' + e.message);
    return [];
  }
}

/**
 * シートのカラム一覧を取得
 * @param {string} sheetName - シート名
 * @returns {Object} カラム情報
 */
function getSheetColumns(sheetName) {
  try {
    var sheet = getSheet(sheetName);
    var headers = sheet.getRange(1, 1, 1, sheet.getLastColumn()).getValues()[0];

    var config = IMPORT_SHEET_CONFIGS[sheetName];

    return {
      success: true,
      columns: headers.filter(function(h) { return h; }),
      idColumn: config ? config.idColumn : null,
      requiredColumns: config ? config.requiredColumns : []
    };
  } catch (e) {
    Logger.log('カラム一覧取得エラー: ' + e.message);
    return { success: false, error: e.message };
  }
}

/**
 * CSVデータを解析
 * @param {string} csvContent - CSV文字列
 * @returns {Object} 解析結果
 */
function parseImportData(csvContent) {
  try {
    var parsed = Utilities.parseCsv(csvContent);

    if (parsed.length < 2) {
      return { success: false, error: 'データが不足しています' };
    }

    var headers = parsed[0];
    var rows = [];

    for (var i = 1; i < parsed.length; i++) {
      var row = {};
      headers.forEach(function(header, index) {
        row[header] = parsed[i][index] || '';
      });
      rows.push(row);
    }

    return {
      success: true,
      headers: headers,
      rows: rows,
      rowCount: rows.length
    };
  } catch (e) {
    Logger.log('CSVパースエラー: ' + e.message);
    return { success: false, error: 'CSVの解析に失敗しました: ' + e.message };
  }
}

/**
 * インポートデータをバリデーション
 * @param {string} sheetName - シート名
 * @param {Array} rows - インポートデータ
 * @param {Object} columnMapping - カラムマッピング（CSVカラム→シートカラム）
 * @returns {Object} バリデーション結果
 */
function validateImportData(sheetName, rows, columnMapping) {
  try {
    var config = IMPORT_SHEET_CONFIGS[sheetName];
    if (!config) {
      return { success: false, error: 'このシートはインポートに対応していません' };
    }

    var errors = [];
    var warnings = [];

    rows.forEach(function(row, index) {
      var rowNum = index + 2; // ヘッダー行 + 1-indexed

      // マッピングを適用
      var mappedRow = {};
      Object.keys(columnMapping).forEach(function(csvCol) {
        var sheetCol = columnMapping[csvCol];
        if (sheetCol) {
          mappedRow[sheetCol] = row[csvCol];
        }
      });

      // 必須項目チェック
      config.requiredColumns.forEach(function(col) {
        if (!mappedRow[col]) {
          errors.push({
            row: rowNum,
            column: col,
            message: col + 'は必須です'
          });
        }
      });

      // カスタムバリデーション
      if (config.validateRow) {
        var result = config.validateRow(mappedRow);
        if (!result.valid) {
          errors.push({
            row: rowNum,
            message: result.error
          });
        }
      }
    });

    return {
      success: true,
      valid: errors.length === 0,
      errors: errors,
      warnings: warnings,
      errorCount: errors.length,
      warningCount: warnings.length
    };
  } catch (e) {
    Logger.log('バリデーションエラー: ' + e.message);
    return { success: false, error: e.message };
  }
}

/**
 * インポートを実行
 * @param {string} sheetName - シート名
 * @param {Array} rows - インポートデータ
 * @param {Object} columnMapping - カラムマッピング
 * @param {Object} options - オプション
 * @returns {Object} 実行結果
 */
function executeImport(sheetName, rows, columnMapping, options) {
  try {
    // 権限チェック
    if (!checkPermission('import.execute')) {
      return { success: false, error: 'インポートの権限がありません' };
    }

    var config = IMPORT_SHEET_CONFIGS[sheetName];
    if (!config) {
      return { success: false, error: 'このシートはインポートに対応していません' };
    }

    var user = getAuditUserInfo();
    var insertedCount = 0;
    var updatedCount = 0;
    var skippedCount = 0;
    var errors = [];

    rows.forEach(function(row, index) {
      try {
        // マッピングを適用
        var mappedRow = {};
        Object.keys(columnMapping).forEach(function(csvCol) {
          var sheetCol = columnMapping[csvCol];
          if (sheetCol) {
            mappedRow[sheetCol] = row[csvCol];
          }
        });

        // IDが存在する場合は更新、なければ新規作成
        var existingId = mappedRow[config.idColumn];

        if (existingId) {
          var existing = findById(sheetName, config.idColumn, existingId);
          if (existing) {
            updateDataWithAudit(sheetName, config.idColumn, existingId, mappedRow);
            updatedCount++;
          } else {
            insertDataWithAudit(sheetName, mappedRow, config.idColumn);
            insertedCount++;
          }
        } else {
          insertDataWithAudit(sheetName, mappedRow, config.idColumn);
          insertedCount++;
        }
      } catch (e) {
        errors.push({
          row: index + 2,
          error: e.message
        });
        skippedCount++;
      }
    });

    // 監査ログ
    AuditLogRepository.log({
      userEmail: user.email,
      userName: user.name,
      department: user.department,
      operationType: 'IMPORT',
      targetSheet: sheetName,
      targetId: 'BULK',
      beforeJson: '',
      afterJson: JSON.stringify({
        insertedCount: insertedCount,
        updatedCount: updatedCount,
        skippedCount: skippedCount
      }),
      changedFields: ''
    });

    return {
      success: true,
      insertedCount: insertedCount,
      updatedCount: updatedCount,
      skippedCount: skippedCount,
      errors: errors
    };
  } catch (e) {
    Logger.log('インポート実行エラー: ' + e.message);
    return { success: false, error: e.message };
  }
}

// ========================================
// 工事台帳 API
// ========================================

/**
 * 工事台帳一覧を取得
 * @param {Object} filter - フィルター条件
 * @returns {Array} 工事台帳一覧
 */
function getLedgers(filter) {
  try {
    var ledgers = LedgerRepository.findAll();

    // フィルター適用
    if (filter) {
      if (filter.projectId) {
        ledgers = ledgers.filter(function(l) {
          return l['案件ID'] === filter.projectId;
        });
      }
      if (filter.status) {
        ledgers = ledgers.filter(function(l) {
          return l['ステータス'] === filter.status;
        });
      }
    }

    // 新しいデータを先頭に表示
    return sanitizeDates(ledgers).reverse();
  } catch (e) {
    Logger.log('工事台帳一覧取得エラー: ' + e.message);
    return [];
  }
}

/**
 * 工事番号で工事台帳を取得（存在しない場合は新規作成用データを返す）
 * @param {string} constructionId - 工事番号
 * @returns {Object} 工事台帳データ（5タブ分のデータを含む）
 */
function getLedgerByConstructionId(constructionId) {
  try {
    // 工事情報を取得
    var construction = ConstructionRepository.findById(constructionId);
    if (!construction) {
      return { success: false, error: '工事が見つかりません: ' + constructionId };
    }

    // 既存の工事台帳を検索
    var ledger = LedgerRepository.findByConstructionId(constructionId);

    // 関連データを集約
    var projectId = construction['関連案件ID'];
    var project = projectId ? ProjectRepository.findById(projectId) : null;

    // 契約協議書を取得
    var contractId = construction['関連契約協議書ID'];
    var contract = null;
    if (contractId) {
      contract = ContractRepository.findById(contractId);
    }

    // 発注業者一覧を取得
    var rawOrders = OrderRepository.findByConstructionId(constructionId);

    // 協力会社マスタのルックアップマップを構築
    var vendorMap = {};
    try {
      var allVendors = VendorRepository.findAll();
      allVendors.forEach(function(v) {
        vendorMap[v['仕入先コード']] = v;
      });
    } catch (e) {
      Logger.log('協力会社マスタ取得エラー: ' + e.message);
    }

    // フロントエンド用にフィールド名をマッピング
    var orders = rawOrders.map(function(order) {
      var mapped = Object.assign({}, order);
      // 施主名(注文決定業者名)列には仕入先コードが格納されている
      var vendorId = order['施主名(注文決定業者名)'] || order['協力会社ID'] || order['仕入先コード'] || '';
      var vendor = vendorMap[vendorId];
      mapped['仕入先コード'] = vendorId;
      mapped['協力会社名'] = vendor ? (vendor['会社名'] || vendor['略称'] || vendorId) : (order['協力会社名'] || vendorId);
      mapped['仕入先名称'] = mapped['協力会社名'];
      mapped['工事内容'] = order['工事内容'] || order['注文内容'] || '';
      mapped['発注金額'] = order['発注金額'] || order['注文金額'] || 0;
      mapped['ステータス'] = order['ステータス'] || order['支払ステータス(協力会社への)'] || '進行中';
      return mapped;
    });

    // 支払記録を取得
    var payments = PaymentRepository.findByProjectId(projectId) || [];
    // この工事に関連する支払のみをフィルタ
    payments = payments.filter(function(p) {
      return orders.some(function(o) {
        return o['発注ID'] === p['発注ID'];
      });
    });

    // 添付ファイルを取得
    var attachments = AttachmentRepository.findByEntity('construction', constructionId);

    // 担当社員を取得
    var employees = [];
    var employeeIds = (construction['担当社員ID'] || '').split(',').filter(function(id) { return id.trim(); });
    employeeIds.forEach(function(empId) {
      var emp = EmployeeRepository.findById(empId.trim());
      if (emp) employees.push(emp);
    });

    // 協力会社情報を取得（ユニークなベンダー一覧）
    var vendors = [];
    orders.forEach(function(order) {
      var vendorId = order['仕入先コード'];
      if (vendorId && vendorMap[vendorId]) {
        var vendor = vendorMap[vendorId];
        if (!vendors.some(function(v) { return v['仕入先コード'] === vendor['仕入先コード']; })) {
          vendors.push(vendor);
        }
      }
    });

    // 顧客情報を取得
    var customer = null;
    if (project && project['顧客ID']) {
      customer = CustomerRepository.findById(project['顧客ID']);
    }

    // フロントエンド向けにフィールド名を補完
    var mappedConstruction = Object.assign({}, construction);
    mappedConstruction['着工日'] = mappedConstruction['着工日'] || formatDateForApi(construction['工期始']) || '';
    mappedConstruction['完工日'] = mappedConstruction['完工日'] || formatDateForApi(construction['工期終']) || '';
    mappedConstruction['工事名称'] = mappedConstruction['工事名称'] || construction['工事名'] || '';
    mappedConstruction['進捗率'] = Number(construction['進捗率']) || 0;

    var mappedProject = project ? Object.assign({}, project) : null;
    if (mappedProject) {
      mappedProject['受注金額'] = mappedProject['受注金額'] || mappedProject['契約金額'] || 0;
    }

    var mappedContract = contract ? Object.assign({}, contract) : null;
    if (mappedContract) {
      mappedContract['請負金額'] = mappedContract['請負金額'] || mappedContract['契約金額'] || 0;
    }

    var result = {
      success: true,
      ledger: ledger || {
        '工事番号': constructionId,
        '案件ID': projectId,
        'ステータス': '作成中'
      },
      isNew: !ledger,

      // タブ1: 工事概要
      overview: {
        construction: mappedConstruction,
        project: mappedProject,
        contract: mappedContract,
        customer: customer
      },

      // タブ2: 元請体制
      staff: {
        employees: employees
      },

      // タブ3: 下請一覧
      subcontractors: {
        orders: orders,
        vendors: vendors
      },

      // タブ4: 支払記録
      payments: {
        payments: payments,
        summary: {
          totalAmount: payments.reduce(function(sum, p) { return sum + (Number(p['支払金額']) || 0); }, 0),
          paidCount: payments.filter(function(p) { return p['支払ステータス'] === '支払済'; }).length,
          pendingCount: payments.filter(function(p) { return p['支払ステータス'] !== '支払済'; }).length
        }
      },

      // タブ5: 完成記録
      completion: {
        progress: construction['進捗率'] || 0,
        progressMemo: construction['進捗メモ'] || '',
        attachments: attachments
      }
    };
    return sanitizeDates(result);
  } catch (e) {
    Logger.log('工事台帳取得エラー: ' + e.message);
    return { success: false, error: e.message };
  }
}

/**
 * 工事台帳を保存
 * @param {Object} ledgerData - 工事台帳データ
 * @returns {Object} 保存結果
 */
function saveLedger(ledgerData) {
  try {
    var result = LedgerRepository.save(ledgerData);
    return { success: true, ledger: result };
  } catch (e) {
    Logger.log('工事台帳保存エラー: ' + e.message);
    return { success: false, error: e.message };
  }
}

/**
 * 工事台帳のステータスを更新
 * @param {string} ledgerId - 台帳ID
 * @param {string} status - 新しいステータス
 * @returns {Object} 更新結果
 */
function updateLedgerStatus(ledgerId, status) {
  try {
    var ledger = LedgerRepository.findById(ledgerId);
    if (!ledger) {
      return { success: false, error: '工事台帳が見つかりません' };
    }

    ledger['ステータス'] = status;
    var result = LedgerRepository.save(ledger);
    return { success: true, ledger: result };
  } catch (e) {
    Logger.log('工事台帳ステータス更新エラー: ' + e.message);
    return { success: false, error: e.message };
  }
}

// ========================================
// マイルストーン API
// ========================================

/**
 * 工事のマイルストーン一覧を取得
 * @param {string} constructionId - 工事番号
 * @returns {Array} マイルストーン一覧
 */
function getMilestones(constructionId) {
  try {
    return sanitizeDates(MilestoneRepository.findByConstructionId(constructionId));
  } catch (e) {
    Logger.log('マイルストーン取得エラー: ' + e.message);
    return [];
  }
}

/**
 * 期間内のマイルストーンを取得（カレンダー/ガントチャート用）
 * @param {string} startDate - 開始日（YYYY-MM-DD）
 * @param {string} endDate - 終了日（YYYY-MM-DD）
 * @returns {Array} マイルストーン一覧（工事情報付き）
 */
function getMilestonesByDateRange(startDate, endDate) {
  try {
    var milestones = MilestoneRepository.findByDateRange(startDate, endDate);

    // 工事情報を付加
    var result = milestones.map(function(ms) {
      var construction = ConstructionRepository.findById(ms['工事番号']);
      return {
        milestone: ms,
        construction: construction,
        projectId: construction ? construction['関連案件ID'] : null
      };
    });
    return sanitizeDates(result);
  } catch (e) {
    Logger.log('マイルストーン期間取得エラー: ' + e.message);
    return [];
  }
}

/**
 * マイルストーンを保存
 * @param {Object} milestoneData - マイルストーンデータ
 * @returns {Object} 保存結果
 */
function saveMilestone(milestoneData) {
  try {
    var result = MilestoneRepository.save(milestoneData);
    return { success: true, milestone: sanitizeDates(result) };
  } catch (e) {
    Logger.log('マイルストーン保存エラー: ' + e.message);
    return { success: false, error: e.message };
  }
}

/**
 * マイルストーンを削除
 * @param {string} milestoneId - マイルストーンID
 * @returns {Object} 削除結果
 */
function deleteMilestone(milestoneId) {
  try {
    MilestoneRepository.delete(milestoneId);
    return { success: true };
  } catch (e) {
    Logger.log('マイルストーン削除エラー: ' + e.message);
    return { success: false, error: e.message };
  }
}

/**
 * 遅延マイルストーンを取得
 * @returns {Array} 遅延マイルストーン一覧
 */
function getDelayedMilestones() {
  try {
    var milestones = MilestoneRepository.findDelayed();

    // 工事情報を付加
    var result = milestones.map(function(ms) {
      var construction = ConstructionRepository.findById(ms['工事番号']);
      var project = construction ? ProjectRepository.findById(construction['関連案件ID']) : null;
      return {
        milestone: ms,
        construction: construction,
        project: project
      };
    });
    return sanitizeDates(result);
  } catch (e) {
    Logger.log('遅延マイルストーン取得エラー: ' + e.message);
    return [];
  }
}

/**
 * 工事の進捗率を更新
 * @param {string} constructionId - 工事番号
 * @param {number} progressRate - 進捗率（0-100）
 * @param {string} memo - 進捗メモ
 * @returns {Object} 更新結果
 */
function updateConstructionProgress(constructionId, progressRate, memo) {
  try {
    var construction = ConstructionRepository.findById(constructionId);
    if (!construction) {
      return { success: false, error: '工事が見つかりません' };
    }

    construction['進捗率'] = progressRate;
    construction['進捗更新日'] = Utilities.formatDate(new Date(), Session.getScriptTimeZone(), 'yyyy-MM-dd');
    construction['進捗メモ'] = memo || '';

    var result = ConstructionRepository.save(construction);
    return { success: true, construction: result };
  } catch (e) {
    Logger.log('工事進捗更新エラー: ' + e.message);
    return { success: false, error: e.message };
  }
}

// ========================================
// スケジュール API
// ========================================

/**
 * ガントチャート用データを取得
 * @param {Object} filter - フィルター条件
 * @returns {Object} ガントチャートデータ
 */
function getGanttChartData(filter) {
  try {
    filter = filter || {};

    // 工事一覧を取得
    var constructions = ConstructionRepository.findAll();

    // フィルター適用
    if (filter.projectId) {
      constructions = constructions.filter(function(c) {
        return c['関連案件ID'] === filter.projectId;
      });
    }
    if (filter.status) {
      constructions = constructions.filter(function(c) {
        return c['ステータス'] === filter.status;
      });
    }

    // 各工事にマイルストーンを付加
    var items = constructions.map(function(construction) {
      var milestones = MilestoneRepository.findByConstructionId(construction['工事番号']);
      var project = ProjectRepository.findById(construction['関連案件ID']);

      return {
        id: construction['工事番号'],
        name: construction['工事名'] || construction['工事名称'] || construction['工事番号'],
        startDate: formatDateForApi(construction['工期始']) || formatDateForApi(construction['着工日']) || '',
        endDate: formatDateForApi(construction['工期終']) || formatDateForApi(construction['完工日']) || '',
        progress: Number(construction['進捗率']) || 0,
        status: construction['進捗ステータス'] || construction['ステータス'] || '',
        project: project ? {
          id: project['案件ID'],
          name: project['案件名']
        } : null,
        milestones: milestones.map(function(ms) {
          return {
            id: ms['マイルストーンID'],
            name: ms['名称'],
            planDate: formatDateForApi(ms['予定日']),
            actualDate: formatDateForApi(ms['実績日']),
            status: ms['ステータス']
          };
        })
      };
    });

    // 開始日でソート
    items.sort(function(a, b) {
      var dateA = new Date(a.startDate || '9999-12-31');
      var dateB = new Date(b.startDate || '9999-12-31');
      return dateA - dateB;
    });

    return {
      success: true,
      items: items
    };
  } catch (e) {
    Logger.log('ガントチャートデータ取得エラー: ' + e.message);
    return { success: false, error: e.message };
  }
}

/**
 * カレンダー用イベントデータを取得
 * @param {string} yearMonth - 年月（YYYY-MM）
 * @returns {Object} カレンダーイベントデータ
 */
function getCalendarEvents(yearMonth) {
  try {
    // 月の開始日と終了日を計算
    var parts = yearMonth.split('-');
    var year = parseInt(parts[0]);
    var month = parseInt(parts[1]);
    var startDate = yearMonth + '-01';
    var lastDay = new Date(year, month, 0).getDate();
    var endDate = yearMonth + '-' + String(lastDay).padStart(2, '0');

    // マイルストーンを取得
    var milestones = MilestoneRepository.findByDateRange(startDate, endDate);

    // 工事の着工日・完工日をイベントとして追加
    var constructions = ConstructionRepository.findAll();
    var events = [];

    // マイルストーンイベント
    milestones.forEach(function(ms) {
      var construction = ConstructionRepository.findById(ms['工事番号']);
      events.push({
        id: ms['マイルストーンID'],
        type: 'milestone',
        date: formatDateForApi(ms['実績日']) || formatDateForApi(ms['予定日']),
        title: ms['名称'],
        status: ms['ステータス'],
        constructionId: ms['工事番号'],
        constructionName: construction ? (construction['工事名'] || construction['工事名称'] || '') : ''
      });
    });

    // 工事開始/完了イベント
    constructions.forEach(function(c) {
      var startDt = formatDateForApi(c['工期始']) || formatDateForApi(c['着工日']) || '';
      var endDt = formatDateForApi(c['工期終']) || formatDateForApi(c['完工日']) || '';
      var cName = c['工事名'] || c['工事名称'] || c['工事番号'];

      if (startDt && startDt >= startDate && startDt <= endDate) {
        events.push({
          id: c['工事番号'] + '-start',
          type: 'construction-start',
          date: startDt,
          title: '着工: ' + cName,
          constructionId: c['工事番号'],
          constructionName: cName
        });
      }

      if (endDt && endDt >= startDate && endDt <= endDate) {
        events.push({
          id: c['工事番号'] + '-end',
          type: 'construction-end',
          date: endDt,
          title: '完工: ' + cName,
          constructionId: c['工事番号'],
          constructionName: cName
        });
      }
    });

    // 日付でソート
    events.sort(function(a, b) {
      return a.date.localeCompare(b.date);
    });

    return {
      success: true,
      yearMonth: yearMonth,
      events: events
    };
  } catch (e) {
    Logger.log('カレンダーイベント取得エラー: ' + e.message);
    return { success: false, error: e.message };
  }
}
