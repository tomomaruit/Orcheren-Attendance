/**
 * スプレッドシートを開いたときにカスタムメニューを追加
 */
function onOpen() {
  SpreadsheetApp.getUi()
      .createMenu('回答整理')
      .addItem('最新の回答をパート順に整理', 'organizeResponses')
      .addSeparator() // 仕切り線
      .addItem('【初回1回だけやる】自動実行トリガーを設定', 'setupTrigger')
      .addToUi();
}

/**
 * この関数を実行すると、「フォーム送信時」に organizeResponses を実行するトリガーが設定される．
 * 重複して設定されないように、既存のトリガーは一旦削除してから作成．
 * この関数は、メニューから一度だけ実行する．
 */
function setupTrigger() {
  const functionToTrigger = 'organizeResponses'; // 自動実行したい関数名

  // 現在のプロジェクトに設定されているトリガーをすべて取得
  const triggers = ScriptApp.getProjectTriggers();
  
  // 同じ関数を対象とする既存のトリガーがあれば削除
  for (const trigger of triggers) {
    if (trigger.getHandlerFunction() === functionToTrigger) {
      ScriptApp.deleteTrigger(trigger);
    }
  }

  // 新しいトリガーを作成
  ScriptApp.newTrigger(functionToTrigger)
    .forSpreadsheet(SpreadsheetApp.getActiveSpreadsheet()) // このスプレッドシートを対象に
    .onFormSubmit() // フォーム送信時に
    .create(); // トリガーを作成

  // ユーザーに完了を通知
  SpreadsheetApp.getUi().alert('フォーム送信時に回答を自動整理する設定が完了しました。');
}

/**
 * フォームの回答を整理するメインの関数
 */
function organizeResponses() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sourceSheetName = 'フォームの回答 1'; // 元データのシート名
  const outputSheetName = '整理済み回答';   // 出力先のシート名

  // 元データを取得
  const sourceSheet = ss.getSheetByName(sourceSheetName);
  if (!sourceSheet) {
    SpreadsheetApp.getUi().alert(`エラー: シート「${sourceSheetName}」が見つかりません。`);
    return;
  }
  const dataRange = sourceSheet.getRange('A2:F199');
  const values = dataRange.getValues();

  // 名前の重複を排除し、最新の回答のみを保持
  const latestResponses = {};
  for (const row of values) {
    const timestamp = row[0]; // A列: タイムスタンプ
    const name = row[1];      // B列: 名前

    if (!name) {
      continue;
    }
    // 同じ名前の回答がないか、または既存の回答よりタイムスタンプが新しければ、データを丸ごと上書き
    if (!latestResponses[name] || latestResponses[name][0] < timestamp) {
      latestResponses[name] = row;
    }
  }

  // オブジェクトから配列に変換
  const uniqueData = Object.values(latestResponses);

  // 指定されたパート順に並び替え
  const partOrder = ['Fl', 'Ob', 'Cl', 'Fg', 'Tp', 'Tb', 'Hr', 'Tu', 'Per', 'Vn', 'Vla', 'Vc', 'Cb']; // ここを任意に並び替え可能
  // ここに入っていない特殊楽器は一番最後になる
  uniqueData.sort((a, b) => {
    const partA = a[2];
    const partB = b[2];
    let indexA = partOrder.indexOf(partA);
    let indexB = partOrder.indexOf(partB);
    if (indexA === -1) indexA = partOrder.length;
    if (indexB === -1) indexB = partOrder.length;
    return indexA - indexB;
  });

  // 新しいシートに結果を書き出し
  let outputSheet = ss.getSheetByName(outputSheetName);
  if (outputSheet) {
    outputSheet.clear();
  } else {
    outputSheet = ss.insertSheet(outputSheetName);
  }

  sourceSheet.getRange('A1:F1').copyTo(outputSheet.getRange('A1:F1'));

  if (uniqueData.length > 0) {
    // 整理したデータを書き込み
    outputSheet.getRange(2, 1, uniqueData.length, uniqueData[0].length).setValues(uniqueData);
    // A列の表示形式を「日時」に設定
    outputSheet.getRange(2, 1, uniqueData.length, 1).setNumberFormat('yyyy/mm/dd hh:mm:ss');
  }
  
  // A列からF列まで (6列分) の幅を150ピクセルに設定する
  outputSheet.setColumnWidths(1, 6, 150);

  // 整理したシートを手前にする
  ss.setActiveSheet(outputSheet);
  ss.moveActiveSheet(1);
}