// 予定と実績を抽出するために、毎日２４時に予定と実績の抽出を実行するトリガーを設定する
function setTrigger() {
  /*// トリガーを削除（トリガーの重複設定防止のため。同じ関数に対して複数のトリガーが存在すると、それらが競合して予期しない動作を引き起こす可能性があるから）
  var triggers = ScriptApp.getProjectTriggers();
  for (var i = 0; i < triggers.length; i++) {
    ScriptApp.deleteTrigger(triggers[i]);
    Logger.log('重複するトリガーを削除しました。');
  }
*/
// 現在の日付を取得
  var now = new Date();
  // 今月の1日の日付を作成
  var firstDayOfMonth = new Date(now.getFullYear(), now.getMonth(), 1);
  // 今月の1日の0時0分0秒に設定
  firstDayOfMonth.setHours(0);
  firstDayOfMonth.setMinutes(0);
  firstDayOfMonth.setSeconds(0);
  

  // 予定機能を24時間に１回行うトリガーの設定。
  ScriptApp.newTrigger('plan_recordTimeEntries')
    .timeBased()
    .atHour(0)  // 24時 (夜中)
    .everyDays(1)  // 毎日実行
    .create();

  Logger.log('予定抽出を２４時間毎に繰り返すトリガーが設定されました。');

  //　実績機能を24時間に１回行うトリガーの設定。
   ScriptApp.newTrigger('result_recordTimeEntries')
    .timeBased()
    .atHour(0)  // 24時 (夜中)
    .everyDays(1)  // 毎日実行
    .create();

   ScriptApp.newTrigger('planDays_copyData') // 実行する関数を指定
    .timeBased() // 時間ベースのトリガーを作成
    .at(firstDayOfMonth) // 毎月1日の0時0分0秒に設定
    .create(); // トリガーを作成する

  Logger.log('予定抽出転記（日毎）を毎月１日毎に繰り返すトリガーが設定されました。');
}

 