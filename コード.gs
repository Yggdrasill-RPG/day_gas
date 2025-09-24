function onOpen() {
  SpreadsheetApp.getUi()
      .createMenu('個人結果レポート')
      .addItem('月と曜日を自動入力', 'autoPopulateMonthAndDay')
      .addToUi();
}

function autoPopulateMonthAndDay() {
  const spreadsheet = SpreadsheetApp.getActiveSpreadsheet();
  const sheet = spreadsheet.getSheetByName('シートの名前'); // シート名を正確に入力してください

  if (!sheet) {
    Browser.msgBox('「シートの名前」というシートが見つかりませんでした。シート名を確認してください。');
    return;
  }

  const today = new Date();
  const year = today.getFullYear();
  const month = today.getMonth(); // 0-11
  
  // A2セルに月を入力
  const monthCell = sheet.getRange('A2');
  monthCell.setValue((month + 1) + '月');
  
  // 今月の最終日を取得
  const lastDayOfMonth = new Date(year, month + 1, 0).getDate();
  
  // 日付と曜日の範囲を処理
  for (let i = 4; i <= 34; i++) { // 4行目から34行目まで（日付1から31）
    const day = sheet.getRange(i, 3).getValue(); // C列（日）
    const dayCell = sheet.getRange(i, 2); // B列（曜日）

    if (day && !isNaN(day)) {
      if (day <= lastDayOfMonth) {
        // 月の最終日内の場合、曜日を計算して入力
        const date = new Date(year, month, day);
        const weekday = date.getDay(); // 0:日, 1:月, ..., 6:土
        
        let weekdayText;
        switch (weekday) {
          case 0: weekdayText = '日'; break;
          case 1: weekdayText = '月'; break;
          case 2: weekdayText = '火'; break;
          case 3: weekdayText = '水'; break;
          case 4: weekdayText = '木'; break;
          case 5: weekdayText = '金'; break;
          case 6: weekdayText = '土'; break;
        }
        
        dayCell.setValue(weekdayText);
      } else {
        // 月の最終日を過ぎた場合、ハイフンを入力
        dayCell.setValue('-');
      }
    } else {
        // C列（日）に値がない場合もクリア
        dayCell.setValue('');
    }
  }

  Browser.msgBox('月と曜日の自動入力が完了しました。');
}
