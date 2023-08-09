// カスタムメニューを作成し、置換実行メニュー項目を追加する
function onOpen() {
  SpreadsheetApp.getUi()
    .createMenu('カスタムメニュー')
    .addItem('置換実行', 'replaceValues')
    .addToUi();
}

// 値を入力するためのプロンプトを表示し、入力された値を返す
function showPrompt(message) {
  var ui = SpreadsheetApp.getUi();
  var result = ui.prompt(message, ui.ButtonSet.OK_CANCEL);
  if (result.getSelectedButton() === ui.Button.OK) {
    return result.getResponseText();
  } else {
    return null;
  }
}

// メッセージを表示するアラートを表示する
function showMessage(message) {
  var ui = SpreadsheetApp.getUi();
  ui.alert(message, ui.ButtonSet.OK);
}

// カスタムメニュー「置換実行」が選択された時に実行される関数
// ユーザーから3つの値を入力し、それらの値を使って処理を実行する
function replaceValues() {
  // アクティブなシートを取得する
  var sheet = SpreadsheetApp.getActiveSheet();

  // 検索する値を入力するプロンプトを表示し、入力を受け取る
  var searchValue = showPrompt('Keyを入力してください');
  if (searchValue === null) return; // キャンセルされた場合は処理を終了

  // 置換する値を入力するプロンプトを表示し、入力を受け取る
  var replaceValue1 = showPrompt('置換するPJCDを入力してください');
  if (replaceValue1 === null) return; // キャンセルされた場合は処理を終了

  // 置換する値をもう1つ入力するプロンプトを表示し、入力を受け取る
  var replaceValue2 = showPrompt('置換するPJCD Nameを入力してください');
  if (replaceValue2 === null) return; // キャンセルされた場合は処理を終了

  // Bカラムの値を配列として取得する
  var range = sheet.getRange('B:B');
  var values = range.getValues();

  var found = false; // 値が見つかったかどうかを表すフラグ

  // Bカラム内で検索する
  for (var i = 0; i < values.length; i++) {
    if (values[i][0] === searchValue) {
      var row = i + 1; // 値が見つかった行
      // 同行のDカラムの値を置換する
      sheet.getRange(row, 4).setValue(replaceValue1);
      // 同行のEカラムの値を置換する
      sheet.getRange(row, 5).setValue(replaceValue2);
      // 置換が行われたセルにジャンプする
      sheet.getRange(row, 1).activate();
      found = true; // 値が見つかったことを示す
    }
  }

  // 置換結果を表示する
  if (found) {
    showMessage('置換が成功しました');
  } else {
    showMessage('入力したKeyが存在しませんでした');
  }
}