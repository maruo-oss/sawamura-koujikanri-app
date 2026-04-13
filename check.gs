/**
 * テンプレスプシの実際のセル内容を確認するスクリプト
 * GASエディタで実行 → ログに全セル内容が出力される
 */
function checkTemplateCells() {
  var TEMPLATE_ID = '1DtFYchjR1uIu59yiheup8QhVs9JQXzaGiKMWZWkjMPs';
  var ss = SpreadsheetApp.openById(TEMPLATE_ID);
  var sheet = ss.getSheetByName('小口契約協議');

  if (!sheet) {
    Logger.log('シートが見つかりません');
    return;
  }

  // 1〜55行, A〜AI列(35列)の全セルを読み取り
  var range = sheet.getRange(1, 1, 55, 35);
  var values = range.getValues();

  for (var i = 0; i < values.length; i++) {
    for (var j = 0; j < values[i].length; j++) {
      var val = values[i][j];
      if (val !== '' && val !== null && val !== undefined) {
        var colLetter = String.fromCharCode(65 + (j < 26 ? j : -1)) ;
        if (j >= 26) {
          colLetter = 'A' + String.fromCharCode(65 + j - 26);
        }
        Logger.log('Row ' + (i+1) + ' Col ' + colLetter + (j+1) + ' (' + colLetter + (i+1) + '): ' + val);
      }
    }
  }

  // マージセル情報
  Logger.log('=== マージセル情報 ===');
  var merges = sheet.getRange(1, 1, 55, 35).getMergedRanges();
  for (var m = 0; m < merges.length; m++) {
    Logger.log(merges[m].getA1Notation());
  }
}
