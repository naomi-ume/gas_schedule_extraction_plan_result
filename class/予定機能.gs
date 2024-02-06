function plan_recordTimeEntries(start, end, sheetName){
  startDate = new Date(start);
  endDate = new Date(end);
  startDate.setHours(0,0,0,0);
  endDate.setHours(24,0,0,0)

  const events = CalendarApp.getCalendarById(gCalendarId).getEvents(startDate, endDate); //カレンダーからイベント取得
  const dataToRecord = {};  //区分１と２が同じ行を合算して格納
  for (let i = 0; i < events.length; i++){
    let event = events[i];
    let title = event.getTitle();
    if(title.indexOf(':') === -1 && title.indexOf('移動') === -1){
      continue;
    }
    let startTime = event.getStartTime();
    let endTime = event.getEndTime();
    let duration;
    duration = (endTime - startTime) / (1000*60*60);//イベント時間の算出

    if (title.indexOf(':') >= 0 && title.split(':').length >= 2){
      const titleParts = title.split(':');
       var group1 = titleParts[0]; //区分１
       var group2 = titleParts[1]; //区分２
       var note = titleParts.slice(2).join(':'); //区分２以降を切り取りコロンで結合してnoteに格納
    }
    else if (title.indexOf('移動') >=0){
      var group1 = '移動';
      var group2 = ' ';
      var note = title;//3つ目以降の要素をnoteへ格納
    }
    // 区分2が空の場合、group2を'*'として処理
    if (group2 === '') {
    var group2 = '*';
    }

    // 時間ID配列のMAX値を算出
    //parseInt は、JavaScriptで文字列を整数に変換するための関数。parseInt は文字列を解析し、整数部分だけを抽出する。1日のミリ秒でend-startを割り、整数にしてる。
    var maxTimeId = parseInt((endDate.getTime() - startDate.getTime()) / ( 1000 * 60 * 60 * 24 ));

    // 区分1、区分2をキーとしてデータを累積。:でgroup1,2を連結させてる
    var key = group1 + ':' + group2;
    //キーが dataToRecord オブジェクトに存在しない場合（!dataToRecord[key] が true の場合）、新しいオブジェクトを作成してデータを初期化する
    if (!dataToRecord[key]) {
      dataToRecord[key] = {
        startTime: startTime,
        endTime: endTime,
        group1: group1,
        group2: group2,
        //maxTimeId + 1 個の要素を持つ新しい配列が作成され、その全ての要素が 0 で初期化されます。これは、後で時間の累積データを格納するための配列。この処理により、指定された key に対応するデータが dataToRecord オブジェクト内に存在しない場合、新しいデータオブジェクトが作成され、それが dataToRecord に追加される
        duration: new Array(maxTimeId + 1).fill(0)
      };
    }

    // 時間IDを算出。startTime から startDate までの経過時間をミリ秒で計算し、それを1日のミリ秒数で割って整数部分を取得。これでdayId には startDate から startTime までの日数が整数として格納。
    var dayId = parseInt((startTime.getTime() - startDate.getTime()) / ( 1000 * 60 * 60 * 24 ));
    //dataToRecord オブジェクト内の key に対応するデータオブジェクトの duration 配列の dayId 番目の要素に、duration を加算。これは、startTime から endDate までの期間内の各日ごとに、対応する duration 配列に時間を累積している
    dataToRecord[key].duration[dayId] += duration;
  }

  // データをスプレッドシートに記録。dataToRecord オブジェクト内の各キーに対してループ処理する。dataToRecord オブジェクト内の各キーに対してループ処理する。key は各反復で現在のキーが代入。for...in ループは、オブジェクト内のプロパティに対して反復処理を行うためのループ構造。この構文は、オブジェクトのプロパティ名（キー）を取得するのに最適。
  for (var key in dataToRecord) {
    //各反復で取得された key に対応するデータオブジェクトを data 変数に代入します。これにより、ループ内で data を通じて各データオブジェクトにアクセスできるようになる
    var data = dataToRecord[key];
    
    // group2が空('*')の場合、group1が一致するもの全てスケジュールに加算
    if (data.group2 === '*') {
      //一時的な合計時間の初期化
      var totalDuration = 0;
      //startTimeからstartDateまでの日数を計算
      var dayId = parseInt((data.startTime.getTime() - startDate.getTime()) / ( 1000 * 60 * 60 * 24 ));
      //ループ: 同じgroup1を持ち、group2が'*'でないデータに対して処理
      for (var subKey in dataToRecord) {
        if (dataToRecord[subKey].group1 === data.group1 && dataToRecord[subKey].group2 != '*') {
          // 合計時間に対象データのdurationを加算
          totalDuration += dataToRecord[subKey].duration[dayId];
        }
      }
      //合計時間を元のデータのdurationに加算
      data.duration[dayId] += dataToRecord[key].duration[dayId];
    }
  }
  //gSheetId で指定されたスプレッドシートのIDを使用して、SpreadsheetApp クラスの openById メソッドでスプレッドシートを開き、その後 getSheetByName メソッドで指定されたシート名（gSheetNamePlan）のシートを取得。取得したシートは spreadsheet 変数に代入
  var spreadsheet = SpreadsheetApp.openById(gSheetId).getSheetByName(gSheetNamePlan);
  /* ヘッダー行を設定。getRange メソッドを使用してスプレッドシート上のセルを指定し、そのセルに対して setValue メソッドを使用して値を設定。スプレッドシートの1行目の1列目（A列1行目）に「区分1」という値を設定,1行目の2列目（B列1行目）に「区分2」という値を設定,スプレッドシートの1行目の3列目（C列1行目）に「合計」という値を設定*/
  spreadsheet.getRange(headerCol, headerRawDirection1).setValue('区分1');
  spreadsheet.getRange(headerCol, headerRawDirection2).setValue('区分2');
  spreadsheet.getRange(headerCol, headerRawDirectionTotal).setValue('合計');

  // 書き込み列を取得する
  //書き込み列
  var col;                
  //startDate を基に新しい Date オブジェクト today を作成
  var today = new Date(startDate);
  //maxTimeId 回ループする for ループが開始。このループは、日数（maxTimeId）分だけ繰り返される。
  //ループ変数 i を初期化する。この変数はループ内で使用され、初期値として 0 が代入。ループが実行される条件を指定。i が maxTimeId より小さい場合にループが続く。ループの各反復が終わると、i の値を 1 増加させる。これにより、i の値が次の反復で使用される。
  for(var i = 0; i < maxTimeId; i++)
  {
    //新しい Date オブジェクト day を作成。startDate から i 日進めた日付を day に設定。day の時刻を0時0分0秒０ミリ秒に設定。これにより、日付だけが残る。
    var day = new Date(startDate);
    day.setDate(startDate.getDate()+i);
    day.setHours(0,0,0,0);
    //col の初期値を 4 に設定し、for (;;) {...} で無限ループを開始。このループは条件を指定せずに永遠に続く。
    for (col=4;;col++)
    {
      // ヘッダ行が''の場合は新たにヘッダを追加。spreadsheet から1行目かつ現在の col 列のセルの値が空 ('') かどうかを確認
      if(spreadsheet.getRange(1, col).getValue() === '')
      {
        //空のセルが見つかった場合、その列の後に新しい列を挿入
        spreadsheet.insertColumnAfter(col);
        //挿入された新しい列の1行目に、計算された日付 day を設定
        spreadsheet.getRange(1, col).setValue(day);
        //条件(日付の一致)が満たされたら、break; を使用して無限ループを終了
        break;
      }
      
      // ヘッダ行が開始日になるまで繰り返し
      // 日付を比較する方法はこんな方法しかないの？→formatDate, getDange, getValueの3つを重複で駆使を避け、一つの関数で比較したい。比較だけに新たに変数を作るのも避けたいとの意図。
      //Utilities.formatDate メソッドは、日付を指定された形式にフォーマットするためのメソッド。スプレッドシートの1行目（ヘッダ行）の col 列目のセルから日付を取得し、'JST' タイムゾーンで 'yyyy/MM/dd' 形式にフォーマット。フォーマットされた日付は headerDate に格納。day を'JST' タイムゾーンで 'yyyy/MM/dd' 形式にフォーマットし、その結果を targetDate に格納。フォーマットされた日付が一致するかどうかを比較。もし一致すれば、break; で無限ループを終了。これにより、特定の日付がスプレッドシートのヘッダ行で見つかると、列の探索終了。
      if(Utilities.formatDate(spreadsheet.getRange(1, col).getValue(), 'JST', 'yyyy/MM/dd')
      === Utilities.formatDate(day, 'JST', 'yyyy/MM/dd'))
      {
        break;
      }
    }
    
    // 最終的な集計結果をスプレッドシートに記録
    //dataToRecord オブジェクトの各プロパティに対して反復処理を行う。key には各プロパティのキーが順番に代入。各プロパティに対応する値（data）を取得します。この data には、オブジェクトの各プロパティに関する情報や集計結果が格納。
    for (var key in dataToRecord) {
      var data = dataToRecord[key];

      // 既にデータがある場合は更新、無い場合は追加。row を2から始め、無限ループを行う。新しい行を探し続ける。
      for (var row=2;;row++)
      {
        // データがない場合は追加。スプレッドシートの row 行目の1列目（A列）のセルが空であるかどうかを確認。
        if(spreadsheet.getRange(row, 1).getValue() === '')
        {
          //spreadsheet.appendRow([data.group1, data.group2, data.duration[i]]);
          // 現在の行に1行追加。
          spreadsheet.insertRowAfter(row);
          // 新規書込み。
          //新しい行に data.group1 の値を1列目に書き込み。
          spreadsheet.getRange(row, 1).setValue(data.group1);
          //新しい行に data.group2 の値を2列目に書き込み。
          spreadsheet.getRange(row, 2).setValue(data.group2);
          /*新しい行 (row 行目) の3列目のセルにアクセスし、そのセルに、SUM関数使ってD列の合計を計算する式を設定。
          新しい行を挿入する度に、その行の D 列の合計を計算し、3列目に表示することが目的。
          '=SUM(D': 文字列として SUM 関数の開始部分を表現。
          row: 　　　row の値がここで文字列に組み込まれる。例えば、もし row が 4 ならば、この部分は文字列 '=SUM(D4' 
          ':': 文字列として範囲の区切りを表現
          row: また、row の値がここで文字列に組み込まれる。例えば、もし row が 4 ならば、この部分は文字列 '4)' 
          ')': 文字列として SUM 関数の終了括弧を表現 */
          spreadsheet.getRange(row, 3).setValue('=SUM(D'+row+':'+row+')');
          //スプレッドシート上の新しい行の特定のセルに、data.duration[i] の値を書き込む処理。data.duration 配列の各要素が新しい行の col 列目に順番に書き込まれ、スプレッドシート上にデータが反映。
          spreadsheet.getRange(row, col).setValue(data.duration[i]);
          break;
        }
        //もしスプレッドシートの row 行目に既にデータが存在する場合、かつそのデータが data.group1 と data.group2 と一致する場合、値を更新
        if(spreadsheet.getRange(row, 1).getValue().toString() === data.group1
        && spreadsheet.getRange(row, 2).getValue().toString() === data.group2)
        {
          // 値更新。既存の行の col 列目に data.duration[i] の値を書き込む。
          spreadsheet.getRange(row, col).setValue(data.duration[i]);
          break;
        }
      }
    }
  }
}


