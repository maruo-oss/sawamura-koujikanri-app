/**
 * Repository層 - Spreadsheetデータアクセス
 *
 * 各シートへのCRUD操作を提供
 */

// ========================================
// 期（会計年度）計算
// ========================================

/**
 * 現在の期を算出する
 * 1期=1951年度（1951年10月〜1952年9月）
 * 10月1日に期が切り替わる
 * 例: 2026年2月 → 75期, 2026年10月 → 76期
 */
function getCurrentFiscalPeriod() {
  var now = new Date();
  var year = now.getFullYear();
  var month = now.getMonth() + 1; // 1-12
  // 10月以降は翌年度扱い
  var fiscalYear = month >= 10 ? year + 1 : year;
  return fiscalYear - 1951;
}

// ========================================
// 全角半角正規化ヘルパー
// ========================================

/**
 * 全角半角の揺らぎを吸収する正規化
 * NFKC正規化: 半角ｶﾀｶﾅ→全角カタカナ、全角英数→半角英数、全角スペース→半角スペース
 */
function normalizeWidth(str) {
  if (str === null || str === undefined || str === '') return '';
  try {
    return String(str).normalize('NFKC').toLowerCase();
  } catch (e) {
    return String(str).toLowerCase();
  }
}

/**
 * 法人格を除去して社名の本体部分を取得
 * 例: "株式会社アルファ" → "アルファ"、"アルファ株式会社" → "アルファ"
 */
var COMPANY_PREFIXES = [
  '株式会社', '有限会社', '合同会社', '合資会社', '合名会社',
  '一般社団法人', '一般財団法人', '公益社団法人', '公益財団法人',
  '特定非営利活動法人', '医療法人', '社会福祉法人', '学校法人',
  '(株)', '(有)', '(合)', '（株）', '（有）', '（合）'
];

function stripCompanyType(name) {
  if (!name) return '';
  // NFKC正規化で ㈱→(株)、㈲→(有) 等に統一してから照合
  var s = String(name).normalize('NFKC').trim();
  for (var i = 0; i < COMPANY_PREFIXES.length; i++) {
    var prefix = COMPANY_PREFIXES[i];
    if (s.indexOf(prefix) === 0) {
      return s.substring(prefix.length).trim();
    }
    if (s.indexOf(prefix) === s.length - prefix.length) {
      return s.substring(0, s.length - prefix.length).trim();
    }
  }
  return s;
}

/**
 * 先頭一致検索（法人格除去対応）
 * 全体の先頭 OR 法人格を除いた社名の先頭 にマッチ
 */
function prefixMatch(value, keyword) {
  try {
    if (!value || !keyword) return false;
    var v = String(value).normalize('NFKC').toLowerCase();
    var kw = String(keyword).normalize('NFKC').toLowerCase();
    // 1. 全体の先頭一致（例: "株式会社アルファ" → "株" でヒット）
    if (v.indexOf(kw) === 0) return true;
    // 2. 法人格除去後の社名先頭一致（例: "株式会社アルファ" → "アルファ" でヒット）
    var stripped = stripCompanyType(value).toLowerCase();
    if (stripped.indexOf(kw) === 0) return true;
    // 3. 末尾の法人格部分の先頭一致（例: "△△建設株式会社" → "株式会社" でヒット）
    if (v.indexOf(stripped) === 0 && stripped.length < v.length) {
      var suffix = v.substring(stripped.length);
      if (suffix.indexOf(kw) === 0) return true;
      // 括弧を除去して比較（例: "㈱" → "(株)" → "株" でもヒット）
      var suffixNoBrackets = suffix.replace(/[()（）]/g, '');
      if (suffixNoBrackets.indexOf(kw) === 0) return true;
    }
    // 4. 先頭の法人格（括弧除去）の先頭一致（例: "㈱アルファ" → "株" でヒット）
    if (stripped.length < v.length && v.indexOf(stripped) > 0) {
      var prefix = v.substring(0, v.indexOf(stripped));
      var prefixNoBrackets = prefix.replace(/[()（）]/g, '');
      if (prefixNoBrackets.indexOf(kw) === 0) return true;
    }
    return false;
  } catch (e) {
    return false;
  }
}

// ========================================
// 共通ヘルパー関数
// ========================================

/**
 * シートを取得
 */
function getSheet(sheetName) {
  var ss = getSpreadsheet();
  var sheet = ss.getSheetByName(sheetName);
  if (!sheet) {
    throw new Error('シートが見つかりません: ' + sheetName);
  }
  return sheet;
}

/**
 * シートの全データをオブジェクト配列で取得
 */
function getAllData(sheetName) {
  var sheet = getSheet(sheetName);
  var lastRow = sheet.getLastRow();
  var lastCol = sheet.getLastColumn();

  if (lastRow < 2) return []; // ヘッダーのみ

  var headers = sheet.getRange(1, 1, 1, lastCol).getValues()[0];
  var data = sheet.getRange(2, 1, lastRow - 1, lastCol).getValues();

  return data.map(function(row) {
    var obj = {};
    headers.forEach(function(header, index) {
      obj[header] = row[index];
    });
    return obj;
  });
}

/**
 * ID・コード照合用の正規化
 * 型・前後空白に加え、全角英数字→半角、不可視文字（ゼロ幅・NBSP・全角空白）も吸収する。
 * ※会社名など本文の比較には使わない（全角を含む正当な文字を壊すため）。ID/コード専用。
 * @param {*} value 任意の値
 * @returns {string} 正規化済み文字列
 */
function normalizeId(value) {
  if (value === null || value === undefined) return '';
  return String(value)
    // 全角英数字・記号(U+FF01-FF5E) -> 半角(U+0021-007E)
    .replace(/[！-～]/g, function(ch) {
      return String.fromCharCode(ch.charCodeAt(0) - 0xFEE0);
    })
    // ゼロ幅スペース・ZWJ・BOM等を除去
    .replace(/[\u200B-\u200D\uFEFF]/g, '')
    // 全角空白(U+3000)・NBSP(U+00A0) -> 半角空白（後段のtrimで除去）
    .replace(/[\u3000\u00A0]/g, ' ')
    .trim();
}

/**
 * 条件に一致するデータを検索
 * 文字列として比較することで型の不一致を防ぐ
 */
function findData(sheetName, conditions) {
  var allData = getAllData(sheetName);

  return allData.filter(function(row) {
    return Object.keys(conditions).every(function(key) {
      // ID/コード照合と同じ正規化で比較（型・全角・空白差を吸収）
      return normalizeId(row[key]) === normalizeId(conditions[key]);
    });
  });
}

/**
 * IDでデータを1件取得
 * 文字列として比較することで型の不一致を防ぐ
 */
function findById(sheetName, idColumn, idValue) {
  var allData = getAllData(sheetName);
  var searchValue = normalizeId(idValue);

  for (var i = 0; i < allData.length; i++) {
    if (normalizeId(allData[i][idColumn]) === searchValue) {
      return allData[i];
    }
  }
  return null;
}

/**
 * データを追加
 */
function insertData(sheetName, data) {
  var sheet = getSheet(sheetName);
  var headers = sheet.getRange(1, 1, 1, sheet.getLastColumn()).getValues()[0];

  var rowValues = headers.map(function(header) {
    return data[header] !== undefined ? data[header] : '';
  });

  sheet.appendRow(rowValues);
  SpreadsheetApp.flush();
  return data;
}

/**
 * データを更新
 */
function updateData(sheetName, idColumn, idValue, data) {
  var sheet = getSheet(sheetName);
  var lastRow = sheet.getLastRow();
  var lastCol = sheet.getLastColumn();

  var headers = sheet.getRange(1, 1, 1, lastCol).getValues()[0];
  var idColIndex = headers.indexOf(idColumn) + 1;

  if (idColIndex === 0) {
    throw new Error('IDカラムが見つかりません: ' + idColumn);
  }

  // IDで行を検索（文字列比較で型の不一致を防ぐ）
  var allData = sheet.getRange(2, 1, lastRow - 1, lastCol).getValues();
  var rowIndex = -1;
  var searchValue = String(idValue).trim();

  for (var i = 0; i < allData.length; i++) {
    if (String(allData[i][idColIndex - 1] || '').trim() === searchValue) {
      rowIndex = i + 2; // 1-indexed + ヘッダー行
      break;
    }
  }

  if (rowIndex === -1) {
    throw new Error('データが見つかりません: ' + idValue);
  }

  // データを更新
  var rowValues = headers.map(function(header) {
    return data[header] !== undefined ? data[header] : allData[rowIndex - 2][headers.indexOf(header)];
  });

  sheet.getRange(rowIndex, 1, 1, lastCol).setValues([rowValues]);
  return data;
}

/**
 * IDでデータを削除
 * @param {string} sheetName - シート名
 * @param {string} idColumn - IDカラム名
 * @param {string} idValue - 削除対象のID値
 * @returns {boolean} 削除成功時true
 */
function deleteData(sheetName, idColumn, idValue) {
  var sheet = getSheet(sheetName);
  var lastRow = sheet.getLastRow();
  var lastCol = sheet.getLastColumn();

  if (lastRow < 2) {
    throw new Error('データが存在しません');
  }

  var headers = sheet.getRange(1, 1, 1, lastCol).getValues()[0];
  var idColIndex = headers.indexOf(idColumn) + 1;

  if (idColIndex === 0) {
    throw new Error('IDカラムが見つかりません: ' + idColumn);
  }

  // IDで行を検索
  var allData = sheet.getRange(2, 1, lastRow - 1, lastCol).getValues();
  var rowIndex = -1;
  var searchValue = String(idValue).trim();

  for (var i = 0; i < allData.length; i++) {
    var rowId = String(allData[i][idColIndex - 1] || '').trim();
    if (rowId === searchValue) {
      rowIndex = i + 2; // 1-indexed + ヘッダー行
      break;
    }
  }

  if (rowIndex === -1) {
    throw new Error('データが見つかりません: ' + idValue);
  }

  // 行を削除
  sheet.deleteRow(rowIndex);
  return true;
}

// ========================================
// 論理削除（ソフトデリート）機能
// ========================================

/**
 * 論理削除に対応したシート一覧
 * これらのシートには「削除フラグ」「削除日時」「削除者」カラムが必要
 */
var SOFT_DELETE_SHEETS = [
  '案件管理',
  '契約協議書',
  '工事管理',
  '発注業者管理',
  '請求管理',
  '支払管理',
  '実行予算',
  '添付ファイル管理'
];

/**
 * シートが論理削除に対応しているか確認
 * @param {string} sheetName - シート名
 * @returns {boolean}
 */
function isSoftDeleteEnabled(sheetName) {
  return SOFT_DELETE_SHEETS.indexOf(sheetName) !== -1;
}

/**
 * 削除済みでないデータのみを取得
 * @param {string} sheetName - シート名
 * @returns {Array} 削除されていないデータの配列
 */
function getAllDataExcludeDeleted(sheetName) {
  var allData = getAllData(sheetName);

  // 論理削除対応シートの場合、削除フラグがTRUEのものを除外
  if (isSoftDeleteEnabled(sheetName)) {
    return allData.filter(function(row) {
      return row['削除フラグ'] !== true && row['削除フラグ'] !== 'TRUE' && row['削除フラグ'] !== 1;
    });
  }

  return allData;
}

/**
 * 削除済みデータのみを取得
 * @param {string} sheetName - シート名
 * @returns {Array} 削除されたデータの配列
 */
function getDeletedData(sheetName) {
  var allData = getAllData(sheetName);

  if (!isSoftDeleteEnabled(sheetName)) {
    return [];
  }

  return allData.filter(function(row) {
    return row['削除フラグ'] === true || row['削除フラグ'] === 'TRUE' || row['削除フラグ'] === 1;
  });
}

/**
 * データを論理削除（ソフトデリート）
 * @param {string} sheetName - シート名
 * @param {string} idColumn - IDカラム名
 * @param {string} idValue - 削除対象のID値
 * @returns {boolean} 削除成功時true
 */
function softDeleteData(sheetName, idColumn, idValue) {
  if (!isSoftDeleteEnabled(sheetName)) {
    throw new Error('このシートは論理削除に対応していません: ' + sheetName);
  }

  var user = getAuditUserInfo();
  var now = Utilities.formatDate(new Date(), Session.getScriptTimeZone(), 'yyyy-MM-dd HH:mm:ss');

  // 削除前のデータを取得
  var before = findById(sheetName, idColumn, idValue);
  if (!before) {
    throw new Error('データが見つかりません: ' + idValue);
  }

  // 削除フラグを設定
  var deleteUpdate = Object.assign({}, before, {
    '削除フラグ': true,
    '削除日時': now,
    '削除者': user.email
  });

  updateData(sheetName, idColumn, idValue, deleteUpdate);

  // 監査ログに記録
  AuditLogRepository.log({
    userEmail: user.email,
    userName: user.name,
    department: user.department,
    operationType: 'DELETE',
    targetSheet: sheetName,
    targetId: idValue,
    beforeJson: JSON.stringify(before),
    afterJson: JSON.stringify(deleteUpdate),
    changedFields: '削除フラグ,削除日時,削除者'
  });

  return true;
}

/**
 * 論理削除されたデータを復元
 * @param {string} sheetName - シート名
 * @param {string} idColumn - IDカラム名
 * @param {string} idValue - 復元対象のID値
 * @returns {boolean} 復元成功時true
 */
function restoreData(sheetName, idColumn, idValue) {
  if (!isSoftDeleteEnabled(sheetName)) {
    throw new Error('このシートは論理削除に対応していません: ' + sheetName);
  }

  var user = getAuditUserInfo();

  // 復元前のデータを取得
  var before = findById(sheetName, idColumn, idValue);
  if (!before) {
    throw new Error('データが見つかりません: ' + idValue);
  }

  // 削除フラグをクリア
  var restoreData = Object.assign({}, before, {
    '削除フラグ': '',
    '削除日時': '',
    '削除者': ''
  });

  updateData(sheetName, idColumn, idValue, restoreData);

  // 監査ログに記録
  AuditLogRepository.log({
    userEmail: user.email,
    userName: user.name,
    department: user.department,
    operationType: 'UPDATE',
    targetSheet: sheetName,
    targetId: idValue,
    beforeJson: JSON.stringify(before),
    afterJson: JSON.stringify(restoreData),
    changedFields: '削除フラグ,削除日時,削除者'
  });

  return true;
}

/**
 * データを完全削除（物理削除）
 * 管理者専用機能
 * @param {string} sheetName - シート名
 * @param {string} idColumn - IDカラム名
 * @param {string} idValue - 削除対象のID値
 * @returns {boolean} 削除成功時true
 */
function hardDeleteData(sheetName, idColumn, idValue) {
  var user = getAuditUserInfo();

  // 削除前のデータを取得
  var before = findById(sheetName, idColumn, idValue);

  var result = deleteData(sheetName, idColumn, idValue);

  // 監査ログに記録
  if (before) {
    AuditLogRepository.log({
      userEmail: user.email,
      userName: user.name,
      department: user.department,
      operationType: 'DELETE',
      targetSheet: sheetName,
      targetId: idValue,
      beforeJson: JSON.stringify(before),
      afterJson: '',
      changedFields: ''
    });
  }

  return result;
}

// ========================================
// 案件管理 Repository
// ========================================

var ProjectRepository = {
  SHEET_NAME: '案件管理',
  ID_COLUMN: '案件ID',

  /**
   * 全案件を取得（削除済み除外）
   */
  findAll: function() {
    return getAllDataExcludeDeleted(this.SHEET_NAME);
  },

  /**
   * 削除済み案件を取得
   */
  findDeleted: function() {
    return getDeletedData(this.SHEET_NAME);
  },

  /**
   * 案件IDで取得
   */
  findById: function(projectId) {
    return findById(this.SHEET_NAME, this.ID_COLUMN, projectId);
  },

  /**
   * 条件で検索
   */
  search: function(filter) {
    var allData = this.findAll();

    return allData.filter(function(row) {
      // 部門フィルタ
      if (filter.department && row['振り分けカテゴリー'] !== filter.department) {
        return false;
      }
      // ステータスフィルタ
      if (filter.status && row['案件ステータス'] !== filter.status) {
        return false;
      }
      // キーワード検索
      if (filter.keyword) {
        var keyword = filter.keyword.toLowerCase();
        var matchName = (row['案件名'] || '').toLowerCase().indexOf(keyword) >= 0;
        var matchCustomer = (row['顧客名'] || '').toLowerCase().indexOf(keyword) >= 0;
        if (!matchName && !matchCustomer) {
          return false;
        }
      }
      return true;
    });
  },

  /**
   * 案件を保存（新規/更新）
   */
  save: function(data) {
    if (data[this.ID_COLUMN]) {
      return updateDataWithAudit(this.SHEET_NAME, this.ID_COLUMN, data[this.ID_COLUMN], data);
    } else {
      data[this.ID_COLUMN] = this.generateId();
      return insertDataWithAudit(this.SHEET_NAME, data, this.ID_COLUMN);
    }
  },

  /**
   * 契約協議書IDで案件を検索
   * 契約協議書から自動登録された案件を特定する（重複防止用）
   * @param {string} contractId - 契約協議書ID（協議書ID or 契約協議書ID）
   * @returns {Object|null} 案件データ（見つからない場合はnull）
   */
  findByContractId: function(contractId) {
    var results = findData(this.SHEET_NAME, { '関連契約協議書ID': contractId });
    return results.length > 0 ? results[0] : null;
  },

  /**
   * 案件を論理削除
   */
  softDelete: function(projectId) {
    return softDeleteData(this.SHEET_NAME, this.ID_COLUMN, projectId);
  },

  /**
   * 案件IDを生成
   */
  generateId: function() {
    var allData = getAllData(this.SHEET_NAME); // 削除済み含む全件でID衝突を防止
    var year = new Date().getFullYear();
    var maxSeq = 0;

    allData.forEach(function(row) {
      var id = String(row['案件ID'] || '');
      var match = id.match(/PJ-(\d{4})-(\d+)/);
      if (match && parseInt(match[1]) === year) {
        var seq = parseInt(match[2]);
        if (seq > maxSeq) maxSeq = seq;
      }
    });

    return 'PJ-' + year + '-' + String(maxSeq + 1).padStart(4, '0');
  }
};

// ========================================
// 工事管理 Repository
// ========================================

var ConstructionRepository = {
  SHEET_NAME: '工事管理',
  ID_COLUMN: '工事番号',

  findAll: function() {
    return getAllDataExcludeDeleted(this.SHEET_NAME);
  },

  findDeleted: function() {
    return getDeletedData(this.SHEET_NAME);
  },

  findById: function(constructionId) {
    return findById(this.SHEET_NAME, this.ID_COLUMN, constructionId);
  },

  findByProjectId: function(projectId) {
    // 文字列として比較するため独自実装（削除済み除外）
    var allData = getAllDataExcludeDeleted(this.SHEET_NAME);
    var searchValue = String(projectId).trim();
    return allData.filter(function(row) {
      var rowProjectId = String(row['関連案件ID'] || '').trim();
      return rowProjectId === searchValue;
    });
  },

  /**
   * 契約協議書IDで工事を検索
   * 契約協議書から自動登録された工事を特定する
   * @param {string} contractId - 契約協議書ID（協議書ID or 契約協議書ID）
   * @returns {Object|null} 工事データ（見つからない場合はnull）
   */
  findByContractId: function(contractId) {
    var results = findData(this.SHEET_NAME, { '関連契約協議書ID': contractId });
    return results.length > 0 ? results[0] : null;
  },

  /**
   * 契約協議書IDに紐づく工事を全件取得（カスケード削除用）
   * @param {string} contractId - 契約協議書ID
   * @returns {Array} 工事データの配列
   */
  findAllByContractId: function(contractId) {
    return findData(this.SHEET_NAME, { '関連契約協議書ID': contractId });
  },

  save: function(data) {
    if (data[this.ID_COLUMN]) {
      return updateDataWithAudit(this.SHEET_NAME, this.ID_COLUMN, data[this.ID_COLUMN], data);
    } else {
      data[this.ID_COLUMN] = this.generateId();
      return insertDataWithAudit(this.SHEET_NAME, data, this.ID_COLUMN);
    }
  },

  softDelete: function(constructionId) {
    return softDeleteData(this.SHEET_NAME, this.ID_COLUMN, constructionId);
  },

  generateId: function() {
    var period = getCurrentFiscalPeriod();
    var prefix = String(period);
    var allData = getAllData(this.SHEET_NAME); // 削除済み含む全件でID衝突を防止
    var maxSeq = 0;

    allData.forEach(function(row) {
      var id = String(row['工事番号'] || '');
      // 新形式(75-00001)と旧形式(C-00001)の両方に対応
      var match = id.match(/^(\d{2})-(\d+)$/) || id.match(/^C-(\d+)$/);
      if (match) {
        var seq = parseInt(match[match.length - 1]);
        if (seq > maxSeq) maxSeq = seq;
      }
    });

    return prefix + '-' + String(maxSeq + 1).padStart(5, '0');
  }
};

// ========================================
// 契約協議書 Repository（ソリューション）
// ========================================

var ContractRepository = {
  SHEET_NAME: '契約協議書',
  ID_COLUMN: '協議書ID',

  findAll: function() {
    return getAllDataExcludeDeleted(this.SHEET_NAME);
  },

  findDeleted: function() {
    return getDeletedData(this.SHEET_NAME);
  },

  findById: function(contractId) {
    return findById(this.SHEET_NAME, this.ID_COLUMN, contractId);
  },

  save: function(data) {
    if (data[this.ID_COLUMN]) {
      return updateDataWithAudit(this.SHEET_NAME, this.ID_COLUMN, data[this.ID_COLUMN], data);
    } else {
      data[this.ID_COLUMN] = this.generateId(data['工事区分']);
      data['作成日'] = new Date();
      return insertDataWithAudit(this.SHEET_NAME, data, this.ID_COLUMN);
    }
  },

  /**
   * 契約協議書を論理削除
   */
  softDelete: function(contractId) {
    return softDeleteData(this.SHEET_NAME, this.ID_COLUMN, contractId);
  },

  /**
   * 協議書IDを生成
   * @param {string} constructionCategory - 工事区分（官庁建築/官庁役務/下請工事/建築工事/土木工事）
   * @returns {string} CT-X-YYYY-NNNN（区分あり）または CT-YYYY-NNNN（区分なし）
   */
  generateId: function(constructionCategory) {
    var prefixMap = {
      '官庁建築': 'K',
      '官庁役務': 'E',
      '下請工事': 'S',
      '建築工事': 'A',
      '土木工事': 'D'
    };
    var prefix = prefixMap[constructionCategory] || '';
    var allData = getAllData(this.SHEET_NAME); // 削除済み含む全件でID衝突を防止
    var year = new Date().getFullYear();
    var maxSeq = 0;

    // プレフィックスあり: CT-X-YYYY-NNNN、なし: CT-YYYY-NNNN
    var pattern = prefix
      ? new RegExp('^CT-' + prefix + '-(\\d{4})-(\\d+)$')
      : /^CT-(\d{4})-(\d+)$/;

    allData.forEach(function(row) {
      var id = String(row['協議書ID'] || '');
      var match = id.match(pattern);
      if (match && parseInt(match[1]) === year) {
        // プレフィックスなしパターンの場合、CT-X-YYYY-NNNNを除外
        if (!prefix && /^CT-[A-Z]-/.test(id)) return;
        var seq = parseInt(match[2]);
        if (seq > maxSeq) maxSeq = seq;
      }
    });

    var seqStr = String(maxSeq + 1).padStart(4, '0');
    return prefix
      ? 'CT-' + prefix + '-' + year + '-' + seqStr
      : 'CT-' + year + '-' + seqStr;
  }
};

// ========================================
// 発注業者管理 Repository
// ========================================

var OrderRepository = {
  SHEET_NAME: '発注業者管理',
  ID_COLUMN: '発注ID',

  findAll: function() {
    return getAllDataExcludeDeleted(this.SHEET_NAME);
  },

  findDeleted: function() {
    return getDeletedData(this.SHEET_NAME);
  },

  findById: function(orderId) {
    return findById(this.SHEET_NAME, this.ID_COLUMN, orderId);
  },

  findByConstructionId: function(constructionId) {
    var allData = getAllDataExcludeDeleted(this.SHEET_NAME);
    return allData.filter(function(row) {
      return row['関連工事番号'] === constructionId;
    });
  },

  save: function(data) {
    if (data[this.ID_COLUMN]) {
      return updateDataWithAudit(this.SHEET_NAME, this.ID_COLUMN, data[this.ID_COLUMN], data);
    } else {
      data[this.ID_COLUMN] = this.generateId();
      return insertDataWithAudit(this.SHEET_NAME, data, this.ID_COLUMN);
    }
  },

  softDelete: function(orderId) {
    return softDeleteData(this.SHEET_NAME, this.ID_COLUMN, orderId);
  },

  generateId: function() {
    var allData = getAllData(this.SHEET_NAME); // 削除済み含む全件でID衝突を防止
    var year = new Date().getFullYear();
    var maxSeq = 0;

    allData.forEach(function(row) {
      var id = String(row['発注ID'] || '');
      var match = id.match(/ORD-(\d{4})-(\d+)/);
      if (match && parseInt(match[1]) === year) {
        var seq = parseInt(match[2]);
        if (seq > maxSeq) maxSeq = seq;
      }
    });

    return 'ORD-' + year + '-' + String(maxSeq + 1).padStart(4, '0');
  }
};

// ========================================
// 見積もり詳細 Repository
// ========================================

var QuoteRepository = {
  SHEET_NAME: '見積もり詳細',
  ID_COLUMN: '見積もり詳細ID',

  findAll: function() {
    return getAllData(this.SHEET_NAME);
  },

  findByOrderId: function(orderId) {
    return findData(this.SHEET_NAME, { '発注ID': orderId });
  },

  findByConstructionId: function(constructionId) {
    return findData(this.SHEET_NAME, { '工事番号': constructionId });
  },

  save: function(data) {
    if (data[this.ID_COLUMN]) {
      return updateDataWithAudit(this.SHEET_NAME, this.ID_COLUMN, data[this.ID_COLUMN], data);
    } else {
      data[this.ID_COLUMN] = this.generateId();
      return insertDataWithAudit(this.SHEET_NAME, data, this.ID_COLUMN);
    }
  },

  delete: function(quoteId) {
    return deleteData(this.SHEET_NAME, this.ID_COLUMN, quoteId);
  },

  generateId: function() {
    var allData = this.findAll();
    var maxSeq = 0;

    allData.forEach(function(row) {
      var id = String(row['見積もり詳細ID'] || '');
      var match = id.match(/Q-(\d+)/);
      if (match) {
        var seq = parseInt(match[1]);
        if (seq > maxSeq) maxSeq = seq;
      }
    });

    return 'Q-' + String(maxSeq + 1).padStart(6, '0');
  }
};

// ========================================
// 顧客マスタ Repository
// ========================================

var CustomerRepository = {
  SHEET_NAME: '顧客マスタ',
  ID_COLUMN: '顧客ID',

  findAll: function() {
    return getAllData(this.SHEET_NAME);
  },

  findById: function(customerId) {
    return findById(this.SHEET_NAME, this.ID_COLUMN, customerId);
  },

  search: function(keyword) {
    var allData = this.findAll();
    if (!keyword) return allData;

    return allData.filter(function(row) {
      try {
        return prefixMatch(row['顧客名'], keyword) ||
               prefixMatch(row['ふりがな'], keyword);
      } catch (e) {
        return false;
      }
    });
  },

  save: function(data) {
    if (data[this.ID_COLUMN]) {
      return updateDataWithAudit(this.SHEET_NAME, this.ID_COLUMN, data[this.ID_COLUMN], data);
    } else {
      data[this.ID_COLUMN] = this.generateId();
      return insertDataWithAudit(this.SHEET_NAME, data, this.ID_COLUMN);
    }
  },

  generateId: function() {
    var allData = this.findAll();
    var maxSeq = 0;

    allData.forEach(function(row) {
      var id = String(row['顧客ID'] || '');
      var match = id.match(/CUS-(\d+)/);
      if (match) {
        var seq = parseInt(match[1]);
        if (seq > maxSeq) maxSeq = seq;
      }
    });

    return 'CUS-' + String(maxSeq + 1).padStart(5, '0');
  }
};

// ========================================
// 社員マスタ Repository
// ========================================

var EmployeeRepository = {
  SHEET_NAME: '社員マスタ',
  ID_COLUMN: '社員番号',

  findAll: function() {
    return getAllData(this.SHEET_NAME);
  },

  findById: function(employeeId) {
    return findById(this.SHEET_NAME, this.ID_COLUMN, employeeId);
  },

  findByEmail: function(email) {
    var results = findData(this.SHEET_NAME, { 'メールアドレス': email });
    return results.length > 0 ? results[0] : null;
  },

  findByDepartment: function(department) {
    return findData(this.SHEET_NAME, { '部門': department });
  },

  save: function(data) {
    if (data[this.ID_COLUMN]) {
      return updateDataWithAudit(this.SHEET_NAME, this.ID_COLUMN, data[this.ID_COLUMN], data);
    } else {
      return insertDataWithAudit(this.SHEET_NAME, data, this.ID_COLUMN);
    }
  }
};

// ========================================
// 協力会社マスタ Repository
// ========================================

var VendorRepository = {
  SHEET_NAME: '協力会社マスタ',
  ID_COLUMN: '仕入先コード',

  findAll: function() {
    return getAllData(this.SHEET_NAME);
  },

  findById: function(vendorId) {
    return findById(this.SHEET_NAME, this.ID_COLUMN, vendorId);
  },

  search: function(keyword) {
    var allData = this.findAll();
    if (!keyword) return allData;

    return allData.filter(function(row) {
      try {
        return prefixMatch(row['会社名'], keyword) ||
               prefixMatch(row['略称'], keyword) ||
               prefixMatch(row['カナ'], keyword);
      } catch (e) {
        return false;
      }
    });
  },

  save: function(data) {
    if (data[this.ID_COLUMN]) {
      return updateDataWithAudit(this.SHEET_NAME, this.ID_COLUMN, data[this.ID_COLUMN], data);
    } else {
      return insertDataWithAudit(this.SHEET_NAME, data, this.ID_COLUMN);
    }
  }
};

// ========================================
// 請求管理 Repository
// ========================================

var InvoiceRepository = {
  SHEET_NAME: '請求管理',
  ID_COLUMN: '請求ID',

  findAll: function() {
    return getAllDataExcludeDeleted(this.SHEET_NAME);
  },

  findDeleted: function() {
    return getDeletedData(this.SHEET_NAME);
  },

  findById: function(invoiceId) {
    return findById(this.SHEET_NAME, this.ID_COLUMN, invoiceId);
  },

  findByProjectId: function(projectId) {
    var allData = getAllDataExcludeDeleted(this.SHEET_NAME);
    var searchValue = String(projectId).trim();
    return allData.filter(function(row) {
      var rowProjectId = String(row['案件ID'] || '').trim();
      return rowProjectId === searchValue;
    });
  },

  findByCustomerId: function(customerId) {
    var allData = getAllDataExcludeDeleted(this.SHEET_NAME);
    var searchValue = normalizeId(customerId);
    return allData.filter(function(row) {
      return normalizeId(row['請求先顧客ID']) === searchValue;
    });
  },

  findByStatus: function(status) {
    var allData = getAllDataExcludeDeleted(this.SHEET_NAME);
    return allData.filter(function(row) {
      return row['入金ステータス'] === status;
    });
  },

  save: function(data) {
    if (data[this.ID_COLUMN]) {
      return updateDataWithAudit(this.SHEET_NAME, this.ID_COLUMN, data[this.ID_COLUMN], data);
    } else {
      data[this.ID_COLUMN] = this.generateId();
      return insertDataWithAudit(this.SHEET_NAME, data, this.ID_COLUMN);
    }
  },

  softDelete: function(invoiceId) {
    return softDeleteData(this.SHEET_NAME, this.ID_COLUMN, invoiceId);
  },

  generateId: function() {
    var allData = getAllData(this.SHEET_NAME); // 削除済み含む全件でID衝突を防止
    var year = new Date().getFullYear();
    var maxSeq = 0;

    allData.forEach(function(row) {
      var id = String(row['請求ID'] || '');
      var match = id.match(/INV-(\d{4})-(\d+)/);
      if (match && parseInt(match[1]) === year) {
        var seq = parseInt(match[2]);
        if (seq > maxSeq) maxSeq = seq;
      }
    });

    return 'INV-' + year + '-' + String(maxSeq + 1).padStart(4, '0');
  }
};

// ========================================
// 支払管理 Repository
// ========================================

var PaymentRepository = {
  SHEET_NAME: '支払管理',
  ID_COLUMN: '支払ID',

  findAll: function() {
    return getAllDataExcludeDeleted(this.SHEET_NAME);
  },

  findDeleted: function() {
    return getDeletedData(this.SHEET_NAME);
  },

  findById: function(paymentId) {
    return findById(this.SHEET_NAME, this.ID_COLUMN, paymentId);
  },

  findByOrderId: function(orderId) {
    var allData = getAllDataExcludeDeleted(this.SHEET_NAME);
    var searchValue = String(orderId).trim();
    return allData.filter(function(row) {
      var rowOrderId = String(row['発注ID'] || '').trim();
      return rowOrderId === searchValue;
    });
  },

  findByProjectId: function(projectId) {
    var allData = getAllDataExcludeDeleted(this.SHEET_NAME);
    var searchValue = String(projectId).trim();
    return allData.filter(function(row) {
      var rowProjectId = String(row['案件ID'] || '').trim();
      return rowProjectId === searchValue;
    });
  },

  findByVendorId: function(vendorId) {
    var allData = getAllDataExcludeDeleted(this.SHEET_NAME);
    var searchValue = normalizeId(vendorId);
    return allData.filter(function(row) {
      return normalizeId(row['協力会社ID']) === searchValue;
    });
  },

  findByMonth: function(yearMonth) {
    // yearMonth: 'YYYY-MM' 形式
    var allData = getAllDataExcludeDeleted(this.SHEET_NAME);
    return allData.filter(function(row) {
      var paymentMonth = row['請求月'] || '';
      return String(paymentMonth).indexOf(yearMonth) === 0;
    });
  },

  findByStatus: function(status) {
    var allData = getAllDataExcludeDeleted(this.SHEET_NAME);
    return allData.filter(function(row) {
      return row['支払ステータス'] === status;
    });
  },

  save: function(data) {
    if (data[this.ID_COLUMN]) {
      return updateDataWithAudit(this.SHEET_NAME, this.ID_COLUMN, data[this.ID_COLUMN], data);
    } else {
      data[this.ID_COLUMN] = this.generateId();
      return insertDataWithAudit(this.SHEET_NAME, data, this.ID_COLUMN);
    }
  },

  generateId: function() {
    var allData = getAllData(this.SHEET_NAME); // 削除済み含む全件でID衝突を防止
    var year = new Date().getFullYear();
    var maxSeq = 0;

    allData.forEach(function(row) {
      var id = String(row['支払ID'] || '');
      var match = id.match(/PAY-(\d{4})-(\d+)/);
      if (match && parseInt(match[1]) === year) {
        var seq = parseInt(match[2]);
        if (seq > maxSeq) maxSeq = seq;
      }
    });

    return 'PAY-' + year + '-' + String(maxSeq + 1).padStart(4, '0');
  }
};

// ========================================
// 支払履歴 Repository
// ========================================

var PaymentHistoryRepository = {
  SHEET_NAME: '支払履歴',
  ID_COLUMN: '支払履歴ID',

  findAll: function() {
    return getAllDataExcludeDeleted(this.SHEET_NAME);
  },

  findById: function(historyId) {
    return findById(this.SHEET_NAME, this.ID_COLUMN, historyId);
  },

  findByPaymentId: function(paymentId) {
    var allData = getAllDataExcludeDeleted(this.SHEET_NAME);
    var searchValue = String(paymentId).trim();
    return allData.filter(function(row) {
      return String(row['支払ID'] || '').trim() === searchValue;
    });
  },

  findByOrderId: function(orderId) {
    var allData = getAllDataExcludeDeleted(this.SHEET_NAME);
    var searchValue = String(orderId).trim();
    return allData.filter(function(row) {
      return String(row['発注ID'] || '').trim() === searchValue;
    });
  },

  save: function(data) {
    if (data[this.ID_COLUMN]) {
      return updateDataWithAudit(this.SHEET_NAME, this.ID_COLUMN, data[this.ID_COLUMN], data);
    } else {
      data[this.ID_COLUMN] = this.generateId();
      return insertDataWithAudit(this.SHEET_NAME, data, this.ID_COLUMN);
    }
  },

  generateId: function() {
    var allData = getAllData(this.SHEET_NAME); // 削除済み含む全件でID衝突を防止
    var year = new Date().getFullYear();
    var maxSeq = 0;

    allData.forEach(function(row) {
      var id = String(row['支払履歴ID'] || '');
      var match = id.match(/PHY-(\d{4})-(\d+)/);
      if (match && parseInt(match[1]) === year) {
        var seq = parseInt(match[2]);
        if (seq > maxSeq) maxSeq = seq;
      }
    });

    return 'PHY-' + year + '-' + String(maxSeq + 1).padStart(4, '0');
  }
};

// ========================================
// 入金履歴 Repository
// ========================================

var InvoicePaymentHistoryRepository = {
  SHEET_NAME: '入金履歴',
  ID_COLUMN: '入金履歴ID',

  findAll: function() {
    return getAllDataExcludeDeleted(this.SHEET_NAME);
  },

  findById: function(historyId) {
    return findById(this.SHEET_NAME, this.ID_COLUMN, historyId);
  },

  findByInvoiceId: function(invoiceId) {
    var allData = getAllDataExcludeDeleted(this.SHEET_NAME);
    var searchValue = String(invoiceId).trim();
    return allData.filter(function(row) {
      return String(row['請求ID'] || '').trim() === searchValue;
    });
  },

  save: function(data) {
    if (data[this.ID_COLUMN]) {
      return updateDataWithAudit(this.SHEET_NAME, this.ID_COLUMN, data[this.ID_COLUMN], data);
    } else {
      data[this.ID_COLUMN] = this.generateId();
      return insertDataWithAudit(this.SHEET_NAME, data, this.ID_COLUMN);
    }
  },

  softDelete: function(historyId) {
    return softDeleteData(this.SHEET_NAME, this.ID_COLUMN, historyId);
  },

  generateId: function() {
    var allData = getAllData(this.SHEET_NAME);
    var year = new Date().getFullYear();
    var maxSeq = 0;

    allData.forEach(function(row) {
      var id = String(row['入金履歴ID'] || '');
      var match = id.match(/IHY-(\d{4})-(\d+)/);
      if (match && parseInt(match[1]) === year) {
        var seq = parseInt(match[2]);
        if (seq > maxSeq) maxSeq = seq;
      }
    });

    return 'IHY-' + year + '-' + String(maxSeq + 1).padStart(4, '0');
  }
};

// ========================================
// 実行予算 Repository
// ========================================

var BudgetRepository = {
  SHEET_NAME: '実行予算',
  ID_COLUMN: '予算ID',

  findAll: function() {
    return getAllDataExcludeDeleted(this.SHEET_NAME);
  },

  findDeleted: function() {
    return getDeletedData(this.SHEET_NAME);
  },

  findById: function(budgetId) {
    return findById(this.SHEET_NAME, this.ID_COLUMN, budgetId);
  },

  findByProjectId: function(projectId) {
    var allData = getAllDataExcludeDeleted(this.SHEET_NAME);
    var searchValue = String(projectId).trim();
    return allData.filter(function(row) {
      var rowProjectId = String(row['案件ID'] || '').trim();
      return rowProjectId === searchValue;
    });
  },

  findByWorkId: function(workId) {
    var allData = getAllDataExcludeDeleted(this.SHEET_NAME);
    var searchValue = String(workId).trim();
    return allData.filter(function(row) {
      var rowWorkId = String(row['工事ID'] || '').trim();
      return rowWorkId === searchValue;
    });
  },

  /**
   * 案件サマリーを集計
   * @param {string} projectId - 案件ID
   * @returns {Object} サマリー情報
   */
  getBudgetSummaryByProjectId: function(projectId) {
    var budgets = this.findByProjectId(projectId);
    var summary = {
      totalEstimate: 0,
      totalBudget: 0,
      totalOrdered: 0,
      totalDifference: 0,
      averageReductionRate: 0,
      itemCount: budgets.length
    };

    budgets.forEach(function(row) {
      summary.totalEstimate += Number(row['見積金額(契約時)'] || 0);
      summary.totalBudget += Number(row['実行予算額'] || 0);
      summary.totalOrdered += Number(row['発注済額'] || 0);
    });

    summary.totalDifference = summary.totalBudget - summary.totalOrdered;
    if (summary.totalBudget > 0) {
      summary.averageReductionRate = ((summary.totalBudget - summary.totalOrdered) / summary.totalBudget * 100).toFixed(1);
    }

    return summary;
  },

  save: function(data) {
    data['更新日'] = new Date();

    if (data[this.ID_COLUMN]) {
      return updateDataWithAudit(this.SHEET_NAME, this.ID_COLUMN, data[this.ID_COLUMN], data);
    } else {
      data[this.ID_COLUMN] = this.generateId();
      data['作成日'] = new Date();
      return insertDataWithAudit(this.SHEET_NAME, data, this.ID_COLUMN);
    }
  },

  softDelete: function(budgetId) {
    return softDeleteData(this.SHEET_NAME, this.ID_COLUMN, budgetId);
  },

  generateId: function() {
    var allData = getAllData(this.SHEET_NAME); // 削除済み含む全件でID衝突を防止
    var year = new Date().getFullYear();
    var maxSeq = 0;

    allData.forEach(function(row) {
      var id = String(row['予算ID'] || '');
      var match = id.match(/BDG-(\d{4})-(\d+)/);
      if (match && parseInt(match[1]) === year) {
        var seq = parseInt(match[2]);
        if (seq > maxSeq) maxSeq = seq;
      }
    });

    return 'BDG-' + year + '-' + String(maxSeq + 1).padStart(4, '0');
  }
};

// ========================================
// マスタルックアップ（ID→名前変換）
// ========================================

/**
 * マスタデータのID→名前変換用ルックアップ
 * APIリクエスト内でキャッシュし、重複したマスタ参照を防ぐ
 */
// ========================================
// 監査ログ Repository
// ========================================

var AuditLogRepository = {
  SHEET_NAME: '監査ログ',
  ID_COLUMN: 'ログID',

  /**
   * 全ログを取得（日時降順）
   */
  findAll: function() {
    var data = getAllData(this.SHEET_NAME);
    // タイムスタンプで降順ソート
    return data.sort(function(a, b) {
      var timeA = new Date(a['タイムスタンプ'] || 0);
      var timeB = new Date(b['タイムスタンプ'] || 0);
      return timeB - timeA;
    });
  },

  /**
   * フィルター条件でログを検索
   * @param {Object} filter - 検索条件
   * @param {string} filter.startDate - 開始日（YYYY-MM-DD）
   * @param {string} filter.endDate - 終了日（YYYY-MM-DD）
   * @param {string} filter.userEmail - ユーザーEmail
   * @param {string} filter.operationType - 操作種別
   * @param {string} filter.targetSheet - 対象シート
   * @param {string} filter.targetId - 対象ID
   * @param {number} filter.limit - 取得件数上限
   * @returns {Array} ログ一覧
   */
  find: function(filter) {
    var allData = this.findAll();
    filter = filter || {};

    var filtered = allData.filter(function(row) {
      // 日付範囲フィルター
      if (filter.startDate) {
        var rowDate = String(row['タイムスタンプ'] || '').substring(0, 10);
        if (rowDate < filter.startDate) return false;
      }
      if (filter.endDate) {
        var rowDate = String(row['タイムスタンプ'] || '').substring(0, 10);
        if (rowDate > filter.endDate) return false;
      }

      // ユーザーフィルター
      if (filter.userEmail && row['ユーザーEmail'] !== filter.userEmail) {
        return false;
      }

      // 操作種別フィルター
      if (filter.operationType && row['操作種別'] !== filter.operationType) {
        return false;
      }

      // 対象シートフィルター
      if (filter.targetSheet && row['対象シート'] !== filter.targetSheet) {
        return false;
      }

      // 対象IDフィルター
      if (filter.targetId && row['対象ID'] !== filter.targetId) {
        return false;
      }

      return true;
    });

    // 件数制限
    if (filter.limit && filter.limit > 0) {
      filtered = filtered.slice(0, filter.limit);
    }

    return filtered;
  },

  /**
   * 監査ログを記録
   * @param {Object} logData - ログデータ
   */
  log: function(logData) {
    var data = {
      'ログID': this.generateId(),
      'タイムスタンプ': Utilities.formatDate(new Date(), Session.getScriptTimeZone(), 'yyyy-MM-dd HH:mm:ss'),
      'ユーザーEmail': logData.userEmail || '',
      'ユーザー名': logData.userName || '',
      '部門': logData.department || '',
      '操作種別': logData.operationType || '',
      '対象シート': logData.targetSheet || '',
      '対象ID': logData.targetId || '',
      '変更前JSON': logData.beforeJson || '',
      '変更後JSON': logData.afterJson || '',
      '変更フィールド': logData.changedFields || ''
    };

    try {
      insertData(this.SHEET_NAME, data);
    } catch (e) {
      // 監査ログの保存失敗はログに記録するが、元の操作は止めない
      Logger.log('監査ログ保存エラー: ' + e.message);
    }
  },

  /**
   * ログIDを生成
   */
  generateId: function() {
    var now = new Date();
    var dateStr = Utilities.formatDate(now, Session.getScriptTimeZone(), 'yyyyMMdd-HHmmss');
    var random = Math.floor(Math.random() * 1000).toString().padStart(3, '0');
    return 'LOG-' + dateStr + '-' + random;
  }
};

// ========================================
// 添付ファイル管理 Repository
// ========================================

var AttachmentRepository = {
  SHEET_NAME: '添付ファイル管理',
  ID_COLUMN: 'ファイルID',

  /**
   * 全添付ファイルを取得（削除済み除外）
   */
  findAll: function() {
    return getAllDataExcludeDeleted(this.SHEET_NAME);
  },

  /**
   * 削除済み添付ファイルを取得
   */
  findDeleted: function() {
    return getDeletedData(this.SHEET_NAME);
  },

  /**
   * ファイルIDで取得
   */
  findById: function(fileId) {
    return findById(this.SHEET_NAME, this.ID_COLUMN, fileId);
  },

  /**
   * エンティティに紐づく添付ファイルを取得
   * @param {string} entityType - エンティティ種別（project/construction/contract/order）
   * @param {string} entityId - エンティティID
   * @returns {Array} 添付ファイル一覧
   */
  findByEntity: function(entityType, entityId) {
    var allData = this.findAll();
    return allData.filter(function(row) {
      return row['関連エンティティ種別'] === entityType &&
             row['関連エンティティID'] === entityId;
    });
  },

  /**
   * カテゴリでフィルタリングして取得
   * @param {string} entityType - エンティティ種別
   * @param {string} entityId - エンティティID
   * @param {string} category - カテゴリ
   * @returns {Array} 添付ファイル一覧
   */
  findByEntityAndCategory: function(entityType, entityId, category) {
    var allData = this.findByEntity(entityType, entityId);
    if (!category) return allData;
    return allData.filter(function(row) {
      return row['カテゴリ'] === category;
    });
  },

  /**
   * 添付ファイルを保存
   * @param {Object} data - 添付ファイルデータ
   * @returns {Object} 保存されたデータ
   */
  save: function(data) {
    if (data[this.ID_COLUMN]) {
      return updateDataWithAudit(this.SHEET_NAME, this.ID_COLUMN, data[this.ID_COLUMN], data);
    } else {
      data[this.ID_COLUMN] = this.generateId();
      data['アップロード日時'] = Utilities.formatDate(new Date(), Session.getScriptTimeZone(), 'yyyy-MM-dd HH:mm:ss');
      return insertDataWithAudit(this.SHEET_NAME, data, this.ID_COLUMN);
    }
  },

  /**
   * 添付ファイルを論理削除
   * @param {string} fileId - ファイルID
   * @returns {boolean}
   */
  softDelete: function(fileId) {
    return softDeleteData(this.SHEET_NAME, this.ID_COLUMN, fileId);
  },

  /**
   * ファイルIDを生成
   */
  generateId: function() {
    var now = new Date();
    var dateStr = Utilities.formatDate(now, Session.getScriptTimeZone(), 'yyyyMMdd');
    var random = Math.floor(Math.random() * 1000000).toString().padStart(6, '0');
    return 'ATT-' + dateStr + '-' + random;
  }
};

// ========================================
// 監査機能付きデータ操作関数
// ========================================

/**
 * 現在のユーザー情報を取得（監査用）
 * @returns {Object} ユーザー情報
 */
function getAuditUserInfo() {
  var email = Session.getActiveUser().getEmail();
  var employee = null;

  try {
    employee = EmployeeRepository.findByEmail(email);
  } catch (e) {
    Logger.log('社員マスタ検索エラー: ' + e.message);
  }

  return {
    email: email,
    name: employee ? employee['氏名'] : email.split('@')[0],
    department: employee ? employee['部門'] : ''
  };
}

/**
 * オブジェクトの変更フィールドを検出
 * @param {Object} before - 変更前データ
 * @param {Object} after - 変更後データ
 * @returns {Array} 変更されたフィールド名の配列
 */
function detectChangedFields(before, after) {
  var changed = [];
  var allKeys = Object.keys(Object.assign({}, before, after));

  allKeys.forEach(function(key) {
    var beforeVal = before ? String(before[key] || '') : '';
    var afterVal = after ? String(after[key] || '') : '';
    if (beforeVal !== afterVal) {
      changed.push(key);
    }
  });

  return changed;
}

/**
 * 監査ログ付きでデータを追加
 * @param {string} sheetName - シート名
 * @param {Object} data - 追加するデータ
 * @param {string} idColumn - IDカラム名
 * @returns {Object} 追加されたデータ
 */
function insertDataWithAudit(sheetName, data, idColumn) {
  var result = insertData(sheetName, data);
  var user = getAuditUserInfo();

  AuditLogRepository.log({
    userEmail: user.email,
    userName: user.name,
    department: user.department,
    operationType: 'CREATE',
    targetSheet: sheetName,
    targetId: data[idColumn] || '',
    beforeJson: '',
    afterJson: JSON.stringify(data),
    changedFields: Object.keys(data).join(',')
  });

  return result;
}

/**
 * 監査ログ付きでデータを更新
 * @param {string} sheetName - シート名
 * @param {string} idColumn - IDカラム名
 * @param {string} idValue - ID値
 * @param {Object} data - 更新するデータ
 * @returns {Object} 更新されたデータ
 */
function updateDataWithAudit(sheetName, idColumn, idValue, data) {
  // 変更前のデータを取得
  var before = findById(sheetName, idColumn, idValue);
  var result = updateData(sheetName, idColumn, idValue, data);
  var user = getAuditUserInfo();

  var changedFields = detectChangedFields(before, data);

  AuditLogRepository.log({
    userEmail: user.email,
    userName: user.name,
    department: user.department,
    operationType: 'UPDATE',
    targetSheet: sheetName,
    targetId: idValue,
    beforeJson: before ? JSON.stringify(before) : '',
    afterJson: JSON.stringify(data),
    changedFields: changedFields.join(',')
  });

  return result;
}

/**
 * 監査ログ付きでデータを削除
 * @param {string} sheetName - シート名
 * @param {string} idColumn - IDカラム名
 * @param {string} idValue - ID値
 * @returns {boolean} 削除成功時true
 */
function deleteDataWithAudit(sheetName, idColumn, idValue) {
  // 削除前のデータを取得
  var before = findById(sheetName, idColumn, idValue);
  var result = deleteData(sheetName, idColumn, idValue);
  var user = getAuditUserInfo();

  AuditLogRepository.log({
    userEmail: user.email,
    userName: user.name,
    department: user.department,
    operationType: 'DELETE',
    targetSheet: sheetName,
    targetId: idValue,
    beforeJson: before ? JSON.stringify(before) : '',
    afterJson: '',
    changedFields: ''
  });

  return result;
}

// ========================================
// マスタルックアップ（ID→名前変換）
// ========================================

/**
 * マスタデータのID→名前変換用ルックアップ
 * APIリクエスト内でキャッシュし、重複したマスタ参照を防ぐ
 */
var MasterLookup = {
  _customerMap: null,
  _vendorMap: null,

  /**
   * 顧客マスタのID→名前マップを取得
   * @returns {Object} 顧客ID -> 顧客名のマップ
   */
  getCustomerMap: function() {
    if (!this._customerMap) {
      this._customerMap = {};
      var customers = CustomerRepository.findAll();
      var self = this;
      customers.forEach(function(c) {
        // キーは findById と同じく String().trim() で正規化（型・空白差での照合漏れ防止）
        self._customerMap[normalizeId(c['顧客ID'])] = c['顧客名'];
      });
    }
    return this._customerMap;
  },

  /**
   * 協力会社マスタのID→名前マップを取得
   * @returns {Object} 仕入先コード -> 会社名のマップ
   */
  getVendorMap: function() {
    if (!this._vendorMap) {
      this._vendorMap = {};
      var vendors = VendorRepository.findAll();
      var self = this;
      vendors.forEach(function(v) {
        // キーは findById と同じく String().trim() で正規化（型・空白差での照合漏れ防止）
        self._vendorMap[normalizeId(v['仕入先コード'])] = v['会社名'];
      });
    }
    return this._vendorMap;
  },

  /**
   * 顧客IDから顧客名を取得
   * @param {string} id 顧客ID
   * @returns {string} 顧客名（見つからない場合はIDをそのまま返す）
   */
  getCustomerName: function(id) {
    return this.getCustomerMap()[normalizeId(id)] || id;
  },

  /**
   * 仕入先コードから会社名を取得
   * @param {string} id 仕入先コード
   * @returns {string} 会社名（見つからない場合はIDをそのまま返す）
   */
  getVendorName: function(id) {
    return this.getVendorMap()[normalizeId(id)] || id;
  },

  /**
   * キャッシュをクリア（マスタ更新時に呼び出し）
   */
  clearCache: function() {
    this._customerMap = null;
    this._vendorMap = null;
  }
};

// ========================================
// 工事台帳 Repository
// ========================================

var LedgerRepository = {
  SHEET_NAME: '工事台帳',
  ID_COLUMN: '台帳ID',

  /**
   * 全工事台帳を取得（削除済み除外）
   */
  findAll: function() {
    return getAllDataExcludeDeleted(this.SHEET_NAME);
  },

  /**
   * 削除済み工事台帳を取得
   */
  findDeleted: function() {
    return getDeletedData(this.SHEET_NAME);
  },

  /**
   * 台帳IDで取得
   */
  findById: function(ledgerId) {
    return findById(this.SHEET_NAME, this.ID_COLUMN, ledgerId);
  },

  /**
   * 工事番号で工事台帳を取得
   * @param {string} constructionId - 工事番号
   * @returns {Object|null} 工事台帳データ
   */
  findByConstructionId: function(constructionId) {
    var allData = this.findAll();
    var searchValue = String(constructionId).trim();
    var results = allData.filter(function(row) {
      return String(row['工事番号'] || '').trim() === searchValue;
    });
    return results.length > 0 ? results[0] : null;
  },

  /**
   * 案件IDで工事台帳一覧を取得
   * @param {string} projectId - 案件ID
   * @returns {Array} 工事台帳一覧
   */
  findByProjectId: function(projectId) {
    var allData = this.findAll();
    var searchValue = String(projectId).trim();
    return allData.filter(function(row) {
      return String(row['案件ID'] || '').trim() === searchValue;
    });
  },

  /**
   * ステータスで工事台帳を取得
   * @param {string} status - ステータス（作成中/確認中/確定）
   * @returns {Array} 工事台帳一覧
   */
  findByStatus: function(status) {
    var allData = this.findAll();
    return allData.filter(function(row) {
      return row['ステータス'] === status;
    });
  },

  /**
   * 工事台帳を保存
   * @param {Object} data - 工事台帳データ
   * @returns {Object} 保存されたデータ
   */
  save: function(data) {
    data['更新日時'] = Utilities.formatDate(new Date(), Session.getScriptTimeZone(), 'yyyy-MM-dd HH:mm:ss');

    if (data[this.ID_COLUMN]) {
      return updateDataWithAudit(this.SHEET_NAME, this.ID_COLUMN, data[this.ID_COLUMN], data);
    } else {
      data[this.ID_COLUMN] = this.generateId();
      data['作成日'] = Utilities.formatDate(new Date(), Session.getScriptTimeZone(), 'yyyy-MM-dd');
      data['ステータス'] = data['ステータス'] || '作成中';
      return insertDataWithAudit(this.SHEET_NAME, data, this.ID_COLUMN);
    }
  },

  /**
   * 台帳IDを生成
   */
  generateId: function() {
    var allData = getAllData(this.SHEET_NAME); // 削除済み含む全件でID衝突を防止
    var year = new Date().getFullYear();
    var maxSeq = 0;

    allData.forEach(function(row) {
      var id = String(row['台帳ID'] || '');
      var match = id.match(/LED-(\d{4})-(\d+)/);
      if (match && parseInt(match[1]) === year) {
        var seq = parseInt(match[2]);
        if (seq > maxSeq) maxSeq = seq;
      }
    });

    return 'LED-' + year + '-' + String(maxSeq + 1).padStart(4, '0');
  }
};

// ========================================
// マイルストーン管理 Repository
// ========================================

var MilestoneRepository = {
  SHEET_NAME: 'マイルストーン管理',
  ID_COLUMN: 'マイルストーンID',

  /**
   * 全マイルストーンを取得
   */
  findAll: function() {
    return getAllData(this.SHEET_NAME);
  },

  /**
   * マイルストーンIDで取得
   */
  findById: function(milestoneId) {
    return findById(this.SHEET_NAME, this.ID_COLUMN, milestoneId);
  },

  /**
   * 工事番号でマイルストーン一覧を取得
   * @param {string} constructionId - 工事番号
   * @returns {Array} マイルストーン一覧（予定日順）
   */
  findByConstructionId: function(constructionId) {
    var allData = this.findAll();
    var searchValue = String(constructionId).trim();
    var results = allData.filter(function(row) {
      return String(row['工事番号'] || '').trim() === searchValue;
    });
    // 予定日でソート（Date型対応）
    return results.sort(function(a, b) {
      var rawA = a['予定日'];
      var rawB = b['予定日'];
      var dateA = (rawA instanceof Date) ? rawA : new Date(rawA || '9999-12-31');
      var dateB = (rawB instanceof Date) ? rawB : new Date(rawB || '9999-12-31');
      return dateA - dateB;
    });
  },

  /**
   * 期間内のマイルストーンを取得
   * @param {string} startDate - 開始日（YYYY-MM-DD）
   * @param {string} endDate - 終了日（YYYY-MM-DD）
   * @returns {Array} マイルストーン一覧
   */
  findByDateRange: function(startDate, endDate) {
    var allData = this.findAll();
    return allData.filter(function(row) {
      var rawPlan = row['予定日'];
      var rawActual = row['実績日'];
      // Date型の場合はyyyy-MM-dd形式に変換
      var planDate = (rawPlan instanceof Date) ? Utilities.formatDate(rawPlan, Session.getScriptTimeZone(), 'yyyy-MM-dd') : String(rawPlan || '');
      var actualDate = (rawActual instanceof Date) ? Utilities.formatDate(rawActual, Session.getScriptTimeZone(), 'yyyy-MM-dd') : String(rawActual || '');
      var targetDate = actualDate || planDate;

      if (!targetDate) return false;
      return targetDate >= startDate && targetDate <= endDate;
    });
  },

  /**
   * ステータスでマイルストーンを取得
   * @param {string} status - ステータス（予定/完了/遅延）
   * @returns {Array} マイルストーン一覧
   */
  findByStatus: function(status) {
    var allData = this.findAll();
    return allData.filter(function(row) {
      return row['ステータス'] === status;
    });
  },

  /**
   * 遅延しているマイルストーンを取得
   * @returns {Array} 遅延マイルストーン一覧
   */
  findDelayed: function() {
    var allData = this.findAll();
    var today = Utilities.formatDate(new Date(), Session.getScriptTimeZone(), 'yyyy-MM-dd');

    return allData.filter(function(row) {
      // 完了していない & 予定日が過ぎている
      var rawPlan = row['予定日'];
      var planDate = (rawPlan instanceof Date) ? Utilities.formatDate(rawPlan, Session.getScriptTimeZone(), 'yyyy-MM-dd') : String(rawPlan || '');
      var status = row['ステータス'];
      return status !== '完了' && planDate && planDate < today;
    });
  },

  /**
   * マイルストーンを保存
   * @param {Object} data - マイルストーンデータ
   * @returns {Object} 保存されたデータ
   */
  save: function(data) {
    // ステータス自動判定
    if (!data['ステータス']) {
      if (data['実績日']) {
        data['ステータス'] = '完了';
      } else {
        var today = Utilities.formatDate(new Date(), Session.getScriptTimeZone(), 'yyyy-MM-dd');
        var planDate = String(data['予定日'] || '');
        data['ステータス'] = planDate < today ? '遅延' : '予定';
      }
    }

    if (data[this.ID_COLUMN]) {
      return updateDataWithAudit(this.SHEET_NAME, this.ID_COLUMN, data[this.ID_COLUMN], data);
    } else {
      data[this.ID_COLUMN] = this.generateId();
      return insertDataWithAudit(this.SHEET_NAME, data, this.ID_COLUMN);
    }
  },

  /**
   * マイルストーンを削除
   * @param {string} milestoneId - マイルストーンID
   * @returns {boolean}
   */
  delete: function(milestoneId) {
    return deleteData(this.SHEET_NAME, this.ID_COLUMN, milestoneId);
  },

  /**
   * マイルストーンIDを生成
   */
  generateId: function() {
    var allData = this.findAll();
    var maxSeq = 0;

    allData.forEach(function(row) {
      var id = String(row['マイルストーンID'] || '');
      var match = id.match(/MS-(\d+)/);
      if (match) {
        var seq = parseInt(match[1]);
        if (seq > maxSeq) maxSeq = seq;
      }
    });

    return 'MS-' + String(maxSeq + 1).padStart(6, '0');
  }
};
