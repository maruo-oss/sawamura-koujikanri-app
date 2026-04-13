/**
 * 小口契約協議書テンプレート修正スクリプト（v3）
 *
 * 修正内容:
 * 1. 工期: ラベルセル(U18,W18,AD18,AF18)を元に戻し、値セル(V17,X17,AE17,AG17)にプレースホルダー設定
 * 2. 契約書: 要/不要の表示方法をコード側で対応（テンプレ変更なし）
 * 3. 竣工時の入金時期プレースホルダー追加（M43）
 *
 * 使い方: GASエディタで fixSmallContractTemplateV3() を実行
 */
function fixSmallContractTemplateV3() {
  var TEMPLATE_ID = '1DtFYchjR1uIu59yiheup8QhVs9JQXzaGiKMWZWkjMPs';
  var ss = SpreadsheetApp.openById(TEMPLATE_ID);
  var sheet = ss.getSheetByName('小口契約協議');

  if (!sheet) {
    Logger.log('シート「小口契約協議」が見つかりません');
    return;
  }

  // === 工期開始: ラベルを復元し、値セルにプレースホルダーを設定 ===
  sheet.getRange('U18').setValue('年');    // ラベル復元（月値で上書きしてしまっていた）
  sheet.getRange('W18').setValue('月');    // ラベル復元（日値で上書きしてしまっていた）
  sheet.getRange('V17').setValue('<<[工期開始_月]>>');  // 正しい値セル
  sheet.getRange('X17').setValue('<<[工期開始_日]>>');  // 正しい値セル

  // === 工期終了: ラベルを復元し、値セルにプレースホルダーを設定 ===
  sheet.getRange('AD18').setValue('年');   // ラベル復元（月値で上書きしてしまっていた）
  sheet.getRange('AF18').setValue('月');   // ラベル復元（日値で上書きしてしまっていた）
  sheet.getRange('AE17').setValue('<<[工期終了_月]>>'); // 正しい値セル
  sheet.getRange('AG17').setValue('<<[工期終了_日]>>'); // 正しい値セル

  // === 竣工時: 入金時期のプレースホルダー追加 ===
  sheet.getRange('M43').setValue('<<[入金時期_竣工時]>>');

  SpreadsheetApp.flush();
  Logger.log('テンプレート修正（v3）が完了しました');
  Logger.log('- 工期ラベル(年/月)を復元');
  Logger.log('- 工期値セル(V17,X17,AE17,AG17)にプレースホルダー設定');
  Logger.log('- 竣工時の入金時期プレースホルダー(M43)追加');
}
