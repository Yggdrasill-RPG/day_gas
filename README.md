# 個人結果報告書 自動化スクリプト

このGoogle Apps Script（GAS）は、Googleスプレッドシート上の個人結果報告書のテンプレートに、現在の月と曜日を自動で入力します。31日がない月（例：2月、4月）の場合も、自動でハイフンを挿入します。

### 主な機能

- **月と曜日の自動入力**: スクリプトを実行すると、現在の月と各日の曜日が自動で入力されます。
- **月末処理の自動化**: 31日がない月でも、カレンダーの正確さに合わせて空白のセルにハイフンが挿入されます。
- **カスタムメニュー**: スプレッドシート上に専用のメニューが追加され、手動での実行が簡単に行えます。
- **自動実行設定**: トリガーを設定することで、毎月1日に自動で実行することも可能です。

### 使い方

1.  **スプレッドシートの準備**
    このスクリプトは、以下の構造を持つスプレッドシートを想定しています。
    - A2セルに月を入力
    - B列に曜日、C列に日（1から31）が入力されている

2.  **スクリプトの導入**
    1.  対象のGoogleスプレッドシートを開きます。
    2.  上部メニューから「**拡張機能**」>「**Apps Script**」を選択します。
    3.  既存のコードがあればすべて削除し、`autoPopulateMonthAndDay.js`ファイルの内容をコピー＆ペーストします。
    4.  ファイル名を分かりやすい名前に変更し、保存します。

3.  **スクリプトの実行**
    スクリプトを保存すると、スプレッドシートのメニューバーに新しく「**個人結果レポート**」というメニューが追加されます。
    - **手動で実行する場合**: 「個人結果レポート」>「月と曜日を自動入力」をクリックします。
    - **自動で実行する場合**: Apps Scriptの編集画面で、左メニューの「**トリガー**」から、毎月1日に実行されるトリガーを設定します。

---

# Personal Report Automation Script

This Google Apps Script (GAS) automatically populates a personal report template on a Google Sheet with the current month and the corresponding weekdays. It also handles months with fewer than 31 days (e.g., February, April) by automatically inserting hyphens.

### Key Features

- **Automatic Month and Weekday Entry**: Running the script automatically fills in the current month and the weekday for each day.
- **End-of-Month Handling**: For months without 31 days, hyphens are automatically inserted into the blank cells for calendar accuracy.
- **Custom Menu**: A dedicated menu is added to the spreadsheet, allowing for easy manual execution.
- **Automated Execution**: By setting up a trigger, the script can be configured to run automatically on the first day of every month.

### How to Use

1.  **Prepare Your Spreadsheet**
    This script is designed for a spreadsheet with the following structure:
    - Month to be entered in cell A2.
    - Weekdays in column B and days (1 to 31) in column C.

2.  **Install the Script**
    1.  Open your Google Sheet.
    2.  From the top menu, select "**Extensions**" > "**Apps Script**".
    3.  Delete any existing code and copy & paste the content from the `autoPopulateMonthAndDay.js` file.
    4.  Rename the file to a descriptive name and save it.

3.  **Run the Script**
    After saving the script, a new menu item called "**Personal Report**" will appear on your spreadsheet menu bar.
    - **To run manually**: Click on "Personal Report" > "Auto-fill Month and Day".
    - **To run automatically**: In the Apps Script editor, go to the "**Triggers**" menu on the left and set up a trigger to run the script on the first day of every month.
