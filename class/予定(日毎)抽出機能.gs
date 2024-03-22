// 予定タブから前月の予定を１ヶ月分取得。
//指定された2つのスプレッドシート間でデータをコピーするための関数。fromSheetName: コピー元のシート名toSheetName: コピー先のシート名　終了位置を特定するために、　　　　　　　　　　　　　　　最初の行のデータから空でないセルを検索し、その列の数を終了位置として扱う。
function planDays_copyData(fromSheetName, toSheetName)
{
  var fromSheetName = gSheetNamePlan; // コピー元のシート名を指定
  var toSheetName = gSheetNamePlanDays; // コピー先のシート名を指定
  var fromsheet = SpreadsheetApp.openById(gSheetId).getSheetByName(fromSheetName);//コピー元の予定シートを開く
  var tosheet = SpreadsheetApp.openById(gSheetId).getSheetByName(toSheetName);//コピー先の予定(日毎)シートを開く
  var today = new Date();//現在の日付の取得
  var firstDayOfLastMonth = new Date(today.getFullYear(), today.getMonth() -1,1);//先月の初日を取得
  var lastDayOfLastMonth = new Date(today.getFullYear(), today.getMonth(), 0);//先月の最後の日を取得
  var data = fromsheet.getDataRange().getValues();//コピー元のデータ取得
  var endCol = findEndColumn(data);//終了位置を見つける
  var lastRow = tosheet.getLastRow();//コピー先の最終行を取得

// 先月分の予定をコピー
for (var i = 0; i < data.length; i++) {
  var rowDate = new Date(data[i][0]); // 予定シートの日付列をDate型に変換
  var rowData = data[i].slice(0, endCol);//終了位置までのデータを取得
  tosheet.appendRow(rowData);//既存データの最終行の次の行からコピー先のシートに追加
  Logger.log('前月の予定(日毎)を転記しました。');
}
  }
  function findEndColumn(data) {
  for (var i = 0; i < data[0].length; i++) {
    if (data[0][i] === "") {
      return i;
    }
  }
 

   
  
   
   
   














   /* 
    for(var i=5;;i++)//i を初期値 5 で定義。ループ毎にインクリメント// 予定シートの行ヘッダが空白の場合。行ヘッダが空であるか、あるいはセルの日付が end よりも後の場合、ループを抜ける。同じセルの値を日付として解釈し、その日付が end よりも後の日付であるかどうかを比較。
    if(fromsheet.getRange(1, i).getValue() === '' ||
    Utilities.formatDate(fromsheet.getRange(1, i).getValue(), 'JST', 'yyyy/MM/dd') > Utilities.formatDate(end, 'JST', 'yyyy/MM/dd'))
    {
      break;
    }
    var cellValue = fromsheet.getRange(1, i).getValue();
console.log("Cell value:", cellValue, typeof cellValue);
    //それぞれ予定シート１行目i列目とendCol列目のセルの値を取得してUtilities.formatDate を使用して、それぞれのセルの日付を指定されたフォーマット ('yyyy/MM/dd') で比較可能な形に変換
    if(Utilities.formatDate(fromsheet.getRange(1, i).getValue(), 'JST', 'yyyy/MM/dd') > Utilities.formatDate(fromsheet.getRange(1, endCol).getValue(), 'JST', 'yyyy/MM/dd'))
    {
      //最も後の日付を持つ列の列番号。特定の日付条件を超えた場合、endColの更新
      endCol = fromsheet.getRange(1, i).getValue();
    }
  }
  //コピー元からコピー先へデータをコピー。コピー元のデータの開始列を指定しています。ここでは0なので、コピー元のデータは最初の列から始まります。getMaxRows() は、スプレッドシートのシートで利用可能な最大行数を返すメソッド
  fromsheet.copyValuesToRange(tosheet, 0, endCol, 0, fromsheet.getMaxRows());

  // 特定の範囲の列削除。第1引数 4 は削除を開始する列のインデックス。この場合、4列目から削除が開始。第2引数 endCol - 3 は削除する列の数。endCol はループで特定された終了列の値で、それから3を引いてる。削除される列の数が計算されている
  fromsheet.deleteColumns(4, endCol - 3);
}*/
/*
// wrapper関数の使用　現在の月。addSumCol 関数が gSheetNamePlan シートからデータを取得し、それを gSheetNameResult シートに追加処理を行いながら書き込む
function addSumColWrapper() {
  var today = new Date();
  var month = today.getMonth() + 1;//月は0から始まるため、+1して実際の月に変換
  addSumCol(gSheetNamePlan, gSheetNameResult,month);
}

// 集計結果を実績タブに転記する。addSumColは指定された2つのスプレッドシートからデータを取得し、それを別のスプレッドシートに転記するための関数
function addSumCol(fromSheetName, toSheetName, month)
{
  var fromsheet = SpreadsheetApp.openById(gSheetId).getSheetByName(fromSheetName);//コピー元のシート開く
  var tosheet = SpreadsheetApp.openById(gSheetId).getSheetByName(toSheetName);//コピー先のシート開く
  var today = new Date();//
  
  if (!tosheet) {
  Logger.log('Error: tosheet not found.');

  // ヘッダー行を設定
  tosheet.getRange(1, 1).setValue('区分1');
  tosheet.getRange(1, 2).setValue('区分2');

  // 列を追加
  var row_f;
  var row_t;
  var col_t;//新しい列の追加を行う際に、既存の列の数を調べるために使用されている。col_t は新しい列の位置を示す変数であり、ループを通じて増えていく。
  for(col_t=3;;col_t++)
  {

    var work=tosheet.getRange(1,col_t).getValue();//もしworkが空（値がない）なら、ループを抜ける
    if(tosheet.getRange(1,col_t).getValue() === '')
    {
      break;
    }
  }
  // 後続に列を追加
  tosheet.insertColumnAfter(col_t);
  // 行ヘッダを追加。month という引数から年と月を取得しています。そして、その年と月を用いて新しい Date オブジェクト thismonth を生成しています。1 は日付を示しており、この場合は月初めを指定
  var thismonth = new Date(month.getFullYear(), month.getMonth(), 1);
  Logger.log('month type: ' + typeof month + ', month value: ' + month);
  //startDateの初期化
  startDate.setHours(0,0,0,0);
  //tosheet の1行目、col_t 列目のセルに、thismonth の日付を JSTタイムゾーンでyyyy/MMフォーマットに変換してセット。新しい列のヘッダーが設定
  tosheet.getRange(1, col_t).setValue(Utilities.formatDate(thismonth, 'JST', 'yyyy/MM'));

  // Fromから抽出。row_fはfromsheet の行を指定するための変数。初期値は2で、2行目からループ処理を始める
  for(row_f=2;;row_f++)
  {
    //fromsheet の指定された行 (row_f) の1列目が空白かどうかを確認。もし空白であれば、データの終了を意味し、ループを抜ける
    if(fromsheet.getRange(row_f,1) === '')
    {
      break;
    }

    // toに転記。row_t=2からループ処理。
    for(row_t=2;;row_t++)
    {
      if(tosheet.getRange(row_t,1) === '')//tosheet の指定された行 (row_t) の1列目が空白かどうかを確認。もし空白であれば、新しい行を挿入し、fromsheet の対応する行の1列目と2列目の値を tosheet に転記します。そして、ループを抜ける。
      {
        tosheet.insertRowAfter(row_t);
        tosheet.getRange(row_t, 1).setValue(fromsheet.getRange(row_f, 1));
        tosheet.getRange(row_t, 1).setValue(fromsheet.getRange(row_f, 2));
        break;
      }
    //fromsheet の現在の行の1列目の値が tosheet の対応する行の1列目の値と一致するかどうかを確認。もし一致する場合、それは同じデータがすでに tosheet に存在することを示し、ループを抜ける。
      if(fromsheet.getRange(row_f,1).getValue() === toSheetName.getRange(row_t,1).getValue())
      {
        break;
      }
    }
    //tosheet の対応する行と列に、fromsheet の現在の行の3列目の値を転記している
    tosheet.getRange(row_t, col).setValue(fromsheet.getRange(row_f, 3));
  }

  // Toで対象行検索→今後addsumcolの機能を拡張するかもしれません
}
}
*/ 
  }