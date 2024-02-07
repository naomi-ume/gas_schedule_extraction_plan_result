// グローバル変数
var gCalendarId = 'naomi.sue.tate@gmail.com'; // カレンダーのIDを指定
var gSheetId = '1nsHDtAjP_AH_qABBDOXA9Djbpcz4e-ILHpI6H_79IUk'; // スプレッドシートのIDを指定
var gSheetNameResult = '実績' // 実績の集計結果が入るシート名
var gSheetNamePlan = '予定' // 予定の集計結果が入るシート名
var gSheetNamePlanDays = '予定(日毎)' //予定(日毎)の集計結果が入るシート名
var gSheetNameResultDays = '実績（日毎）' // 実績（日毎）の集計結果が入るシート名
var gSheetNameTabulationPlan = '予定(集計)' //１ヶ月分の予定の、合計値の集計結果が入るシート名
var gSheetNameTabulationResult = '実績(集計)' //１ヶ月分の実績の、合計値の集計結果が入るシート名
var gEditSpan = 7; // 実績編集期間（この日数を過ぎたら集計に反映される）
var gGetSpan = 7; // 予定取得期間（この日数後よりも前の予定を集計する）

//スプレッドシート管理
var headerCol = 1;//スプレッドシートの１列目
var headerRawDirection1 = 1;//スプレッドシート１行目
var headerRawDirection2 = 2;//スプレッドシート２行目
var headerRawDirectionTotal = 3;//スプレッドシート３行目
var startCol = 4;//スプレッドシート４列目