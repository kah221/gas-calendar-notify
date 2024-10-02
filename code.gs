/**
 * トリガーにより自動で実行される関数
 * 処理の流れ
 * - 送信するメッセージを作成
 * - LINENotifyで送信
 */
function executeWithTrigger(){
  // 送信するメッセージを作る
  message = createMessage();
  
  // LINE Notify でメッセージを送信する
  is_successed = sendNotify(message);
  // console.log(is_successed);
}

/**
 * 午前4時から5時の間に自動で実行される関数
 * 処理の流れ
 * - 本日のイベントをgetEventsFromGCalendar()関数オプション付きで呼び出して取得する
 * - 送信するメッセージを作成
 * - LINENotifyで送信
 */
function executeInMorning(){
  todays_event = getEventsFromGCalendar(true); // 朝自動で行われたという識別のため引数にtrueを指定 2次元配列で取得
  summary = '【自動送信】\n■本日の予定一覧■\n\n';
  if(todays_event.length == 0){
    summary += '本日の予定はありません';
  }else{
    todays_event.forEach(function(event){
      if(event[5] == true){ // 終日予定の場合
        summary += '[終日] ' + event[0] + '\n';
      }else{ // 終日予定でない場合
      // 開始時刻を分解
        let st_h = ('0' + event[1].getHours()).slice(-2);
        let st_m = ('0' + event[1].getMinutes()).slice(-2);
        summary += st_h + ':' + st_m + '~ ' + event[0] + '\n';
      }
    });

    // 現在時刻をフォーマット
    var date = new Date();
    var year = date.getFullYear().toString().slice(-2); // 年の末尾2桁を取得
    var month = (date.getMonth() + 1).toString().padStart(2, '0'); // 月を2桁で表示
    var day = date.getDate().toString().padStart(2, '0'); // 日を2桁で表示
    var hours = date.getHours().toString().padStart(2, '0'); // 時を2桁で表示
    var minutes = date.getMinutes().toString().padStart(2, '0'); // 分を2桁で表示
    var now = `${year}${month}${day}_${hours}${minutes}`;

    summary += '\n以上 ' + todays_event.length + ' 件（' + now + '時点）';
  }
  console.log(summary);
  result = sendNotify(summary);
}

/**
 * Googleカレンダーで更新が起きたときに実行される
 * 処理の流れ
 * - カレンダーIDから本日の全てのイベントを取得
 * - GASのトリガー初期化処理
 * - ｽﾌﾟﾚｯﾄﾞｼｰﾄに書き込むために2次元配列に整形
 *   ↑ 未来のイベントならばトリガーを作成
 *   ↑ この関数が実行されたたとき、まもなくイベントが開始するとき、トリガーを作成せずに即時通知を行う
 * - ｽﾌﾟﾚｯﾄﾞｼｰﾄに書込 ← ｽﾌﾟﾚｯﾄﾞｼｰﾄで管理を行うため（通知済みか否かの確認など）
 */
function getEventsFromGCalendar(isAuto = false) {
  // ------------------------------
  // ↓ 本日の全てのイベントを取得・保存（イベント毎にトリガーも作成）
  // ------------------------------
  const ss = SpreadsheetApp.getActiveSpreadsheet(); // 紐づけているスプシオブジェクトを取得

  // 取得する期間を設定
  var startTime = new Date();
  startTime.setHours(0);
  startTime.setMinutes(0);
  startTime.setSeconds(0);
  var endTime = new Date();
  endTime.setDate(endTime.getDate());
  endTime.setHours(23);
  endTime.setMinutes(59);
  endTime.setSeconds(59);

  // カレンダーIDを指定してカレンダーを取得（Calendarオブジェクト）
  // const calendar_obj = CalendarApp.getCalendarById(PropertiesService.getScriptProperties().getProperty("GOOGLE_CALENDAR_ID_H"));
  const calendar_obj = CalendarApp.getCalendarById(PropertiesService.getScriptProperties().getProperty("GOOGLE_CALENDAR_ID_R"));
  // const calendar_obj = CalendarApp.getCalendarById(PropertiesService.getScriptProperties().getProperty("GOOGLE_CALENDAR_ID_M"));
  console.log(calendar_obj);

  // 期間を指定してイベントを取得（CalendarEventオブジェクト）
  var events_obj = calendar_obj.getEvents(startTime, endTime);

  // ------------------------------
  // ↓ トリガー初期化処理
  // ------------------------------
  // 前回までのトリガーを削除 ← 「何分前に通知させるか」も変更可能の仕様のため
  var trigger_list = ScriptApp.getScriptTriggers();
  for(var i=0; i<trigger_list.length; i++){
    var t = trigger_list[i];
    if(t.getHandlerFunction() === "executeWithTrigger"){ // 時間主導型ではないトリガーを削除する
      ScriptApp.deleteTrigger(t);
      Logger.log('前回のトリガーを削除');
    }
  }

  // ------------------------------
  // ↓ 全イベントを2次元配列に整形 ← ｽﾌﾟﾚｯﾄﾞｼｰﾄに書き込むため
  // ↓ 同時に、未来の予定についてはトリガーを設定
  // ------------------------------
  var events_arr = []; // 設定したカレンダーIDに対応する、本日のイベント全てが2次元配列で格納される
  // イベントの数が0のとき、配列の.lengthでエラーになるのを防ぐif
  if(events_obj != []){
    var isExistInstanceEvent = false; // 即時通知対象イベントが1つでも存在するかどうか
    events_obj.forEach(function(event){
      // ------------------------------
      // ↓ イベントの諸情報を変数として取得
      // ------------------------------
      var title = event.getTitle();         // イベントのタイトルを取得
      var start = event.getStartTime();     // イベントの開始時刻を取得 Dateｵﾌﾞｼﾞｪｸﾄ
      var end = event.getEndTime();         // イベントの終了時刻を取得
      var isAllday = event.isAllDayEvent(); // イベントが終日イベントかどうかを取得
      var isNoticed_1 = false;                // 既に通知済みか否か①
      var isNoticed_2 = false;                // 既に通知済みか否か② 
      var trigger_time_1 = '';              // ｽﾌﾟﾚｯﾄﾞｼｰﾄに記述するため空文字で用意（終日イベ）
      var trigger_time_2 = '';              // 2回目の通知

      // ------------------------------
      // 過去のイベント or 未来のイベント or 即時通知
      // ------------------------------
      if(!isAllday){ // 終日予定でない場合のみトリガーを追加
        var now_time = new Date();  // 現在時刻を取得
        if(start < now_time){       // この関数が実行されたときにイベントの開始時刻が既に過去の場合 → トリガーを作らない、既に通知済みか否かをfalseにする
          isNoticed_1 = true;
          isNoticed_2 = true;
        }else{ // まだイベント開始時刻を迎えていない場合←★
          // ｽﾌﾟﾚｯﾄﾞｼｰﾄから何分前に通知させるか「notice」を取得
          var notice_1 = Number(ss.getSheetByName("設定").getRange("B1").getValue()); // 必ず半角数値5~60までの間であるとする
          var notice_2 = Number(ss.getSheetByName("設定").getRange("B2").getValue());
          
          // トリガーの時刻を作成
          var trigger_time_1 = new Date(start); // 直前ではない
          trigger_time_1.setMinutes(trigger_time_1.getMinutes() - notice_1); // トリガーの時刻を開始時刻「start」からnotice_1分だけ早める
          var trigger_time_2 = new Date(start); // 直前
          trigger_time_2.setMinutes(trigger_time_2.getMinutes() - notice_2);

          // ★且つ、現在が「予定開始と通知時刻の間」のとき、トリガーは作成せず、即時通知を行う
          if(trigger_time_2 < now_time){ // 直前の方を見る
            console.log("即時通知対象イベント");
            isExistInstanceEvent = true; // フラグを立てる
              // ここで、G列に対応するisNoticed_1はfalseのままにしておく
          }else{
            // 次回の実行時間を指定（トリガー機能）
            ScriptApp.newTrigger('executeWithTrigger').timeBased().at(trigger_time_1).create();
            ScriptApp.newTrigger('executeWithTrigger').timeBased().at(trigger_time_2).create();
          }
        }
      }else{ // 既にイベント開始時刻を迎えていた場合
        isNoticed_1 = true; // 終日予定の場合予めtrueにしておき通知を行わない。トリガーも作成しない
        isNoticed_2 = true;
      }

      // 二次元配列に格納
      /**
       * 予定のタイトル [string]
       * 予定の開始時刻 [Dateｵﾌﾞｼﾞｪｸﾄ]
       * 予定の終了時刻 [Dateｵﾌﾞｼﾞｪｸﾄ]
       * 終日予定か否か [bool] trueなら終日予定
       * カレンダー識別用文字列 [string] ← 結局使っていない
       * ポインター [bool] false←まだい通知を行っていない。true←すでに通知を行った又は終日予定
       * ★イベントの長さはメッセージ作成時に計算を行うのでここでは書かない
       * */ 
      events_arr.push([title, start, end, trigger_time_1, trigger_time_2, isAllday, '01', isNoticed_1, isNoticed_2]);
    });
  }
  console.log(events_arr);

  // ------------------------------
  // ↓ ｽﾌﾟﾚｯﾄﾞｼｰﾄに書込
  // ------------------------------
  var sh = ss.getSheetByName("本日の予定一覧");  // シート名（下のタブの部分）を指定

  sh.getRange(2, 1, sh.getLastRow(), sh.getLastColumn()).clearContent(); // 前回の予定を削除
    // sh.getLastRow()    ← データが書き込まれている最後の行を取得
    // sh.getLastColumn() ← データが書き込まれている最後の列を取得

  // イベントの数が0ではない場合のみ書込
  if(events_arr.length != 0){
    sh.getRange(2, 1, events_arr.length, events_arr[0].length).setValues(events_arr);
  }

  // ------------------------------
  // ↑ 書込を行った後に、
  // ↓ 即時通知対象イベントがある場合、通知を行う
  //   （ｽﾌﾟﾚｯﾄﾞｼｰﾄG列のfalseをsendNotifyでtrueに書き換える）
  // ------------------------------
  if(isExistInstanceEvent == true){ // フラグを確認
    message = createMessage(true); // 即時通知イベントなので引数にtrueを指定する
    result = sendNotify(message);
    // console.log(result);
  }

  if(isAuto){
    return events_arr; // 本日のイベント全てを2次元配列で返す
  }else{
    return;
  }
}


/**
 * LINEに送信するメッセージ文字列を作成する関数
 * 処理の流れ
 * - ｽﾌﾟﾚｯﾄﾞｼｰﾄからイベントをすべて取得
 * - ↑をループで1行ずつ読み取る
 * - 「G列：通知済みか否か」がfalseのものについてデータを変数に格納しループを出る
 * - 変数と、ｽﾌﾟﾚｯﾄﾞｼｰﾄに定義したフォーマットに従いにメッセージを作成
 * - ｽﾌﾟﾚｯﾄﾞｼｰﾄの「G列：通知済みか否か」をtrueに上書き
 * 注) 最大文字数：トークンの名前で全角7文字を使った場合、残り半角40文字（スマートウォッチの表示領域より）
 * 
 * @param [bool] isInstance = false - 即時通知イベントか否か
 * @returns [string] message - 送信することになるメッセージ
 */
function createMessage(isInstance = false){
  var message = ''; // 戻り値
  
  // ------------------------------
  // ↓ ｽﾌﾟﾚｯﾄﾞｼｰﾄからすべてのイベントを取得
  // ------------------------------
  const ss = SpreadsheetApp.getActiveSpreadsheet();     // 紐づけているスプシオブジェクトを取得
  var sh = ss.getSheetByName("本日の予定一覧");            // シート名（下のタブの部分）を指定
  var notice_1 = Number(ss.getSheetByName("設定").getRange("B1").getValue()); // 何分前か（1次通知）必ず半角数値5~60までの間であるとする
  var notice_2 = Number(ss.getSheetByName("設定").getRange("B2").getValue()); // 何分前か（2次通知）同上
  var events = sh.getRange(2, 1, sh.getLastRow() - 1, sh.getLastColumn()).getValues();  // イベントを2次元配列で取得 相対座標なのでrowは-1
  Logger.log(events);

  // ------------------------------
  // ↓ ループで1行ずつ読み取る
  /**
   * 流れ
   * 1.i行について
   * 2.1次通知が済みか否か（H列）を確認【★】
   *   →true : 3.へ
   *   →false: 1次の通知メッセージを作成し、H列をtrueにする
   * 3.2次通知が済みか否か（I列）を確認【★★】
   *   →true : 次のi+1行へ
   *   →false: 2次の通知メッセージを作成し、I列をtrueにする
   * 
   * 必ず1次→2次の順で通知が行われる
   * ここで、H列は配列の[7]に、I列は配列の[8]に相当する
   */
  // ------------------------------
  for(var i=0; i<events.length; i++){
    event = events[i];

    // 【★】
    if(event[7] == false){
      console.log("1次通知対象イベント");

      // ------------------------------
      // ↓ メッセージに必要な値を整形
      // ------------------------------
      let event_stt = '';
      let event_end = '';
      let event_len_h = '';
      let event_len_m = '';
      let event_len = '';

      // 開始時刻を分解
      let st_h = ('0' + event[1].getHours()).slice(-2);
      let st_m = ('0' + event[1].getMinutes()).slice(-2);

      // 終了時刻を分解
      let en_h = ('0' + event[2].getHours()).slice(-2);
      let en_m = ('0' + event[2].getMinutes()).slice(-2);

      // 開始・終了時刻の整形
      event_stt = st_h + ':' + st_m; // 開始時刻 例）14:37
      event_end = en_h + ':' + en_m; // 終了時刻

      // 予定の長さ計算
      if(('0' + event[1].getDate()).slice(-2) == ('0' + event[2].getDate()).slice(-2)){ // イベントが日付をまたがないとき
        event_len_h = Number(en_h) - Number(st_h) - 1;
        event_len_m = 60 + Number(en_m) - Number(st_m);
        // 繰り上げ
        if(event_len_m > 59){
          event_len_h ++;
          event_len_m -= 60;
        }
      }else{ // イベントが日付をまたぐとき
        event_len_h = 24 - Number(st_h) + Number(en_h);
        event_len_m = 60 - Number(st_m) + Number(en_m);
      }

      // 予定の長さの整形
      if(event_len_h == 0){
        event_len = event_len_m + 'm';
      }else if(event_len_m == 0){
        event_len = event_len_h + 'h';
      }else{
        event_len = event_len_h + 'h' + event_len_m + 'm';
      }

      // ------------------------------
      // ↓ メッセージを作る
      // ------------------------------
      if(isInstance == true){
        message += "【まもなく】\n"; // 12
      }else{
        message += "【" + notice_1 + "分後】\n\n"; // 9
      }
      // スマートウォッチの表示に収まるようにイベントタイトルを15文字に制限する
      if(event[0].length > 15){
        event_title = event[0].slice(0, 15);
      }else{
        event_title = event[0];
      }

      message += event_title + "\n\n"; // 最大全角15文字
      message += " " + event_stt + "~" + event_end + "\n"; // 12
      message += " " + event_len; // 4又は7

      // ------------------------------
      // ↓ ｽﾌﾟﾚｯﾄﾞｼｰﾄの通知済みか否かをfalse→trueにする
      // ------------------------------
      sh.getRange(2 + i, 7, 1, 1).setValue(true); // iはｲﾝﾃﾞｯｸｽ（行番号-1）なのでカラム名も含め+2が必要 7は1次の通知済みか否かの列番号

      break; // 通知済みか否かの部分がfalseの部分が来た時に抜け出す
    }else if(event[8] == false){ // 【★★】
      console.log("2次通知対象イベント");


      // ------------------------------
      // ↓ メッセージに必要な値を整形
      // ------------------------------
      let event_stt = '';
      let event_end = '';
      let event_len_h = '';
      let event_len_m = '';
      let event_len = '';

      // 開始時刻を分解
      let st_h = ('0' + event[1].getHours()).slice(-2);
      let st_m = ('0' + event[1].getMinutes()).slice(-2);

      // 終了時刻を分解
      let en_h = ('0' + event[2].getHours()).slice(-2);
      let en_m = ('0' + event[2].getMinutes()).slice(-2);

      // 開始・終了時刻の整形
      event_stt = st_h + ':' + st_m; // 開始時刻 例）14:37
      event_end = en_h + ':' + en_m; // 終了時刻

      // 予定の長さ計算
      if(('0' + event[1].getDate()).slice(-2) == ('0' + event[2].getDate()).slice(-2)){ // イベントが日付をまたがないとき
        event_len_h = Number(en_h) - Number(st_h) - 1;
        event_len_m = 60 + Number(en_m) - Number(st_m);
        // 繰り上げ
        if(event_len_m > 59){
          event_len_h ++;
          event_len_m -= 60;
        }
      }else{ // イベントが日付をまたぐとき
        event_len_h = 24 - Number(st_h) + Number(en_h);
        event_len_m = 60 - Number(st_m) + Number(en_m);
      }

      // 予定の長さの整形
      if(event_len_h == 0){
        event_len = event_len_m + 'm';
      }else if(event_len_m == 0){
        event_len = event_len_h + 'h';
      }else{
        event_len = event_len_h + 'h' + event_len_m + 'm';
      }

      // ------------------------------
      // ↓ メッセージを作る
      // ------------------------------
      if(isInstance == true){
        message += "【まもなく】\n"; // 12
      }else{
        message += "【" + notice_2 + "分後】\n\n"; // 9
      }
      // スマートウォッチの表示に収まるようにイベントタイトルを15文字に制限する
      if(event[0].length > 15){
        event_title = event[0].slice(0, 15);
      }else{
        event_title = event[0];
      }

      message += event_title + "\n\n"; // 最大全角15文字
      message += " " + event_stt + "~" + event_end + "\n"; // 12
      message += " " + event_len; // 4又は7

      // ------------------------------
      // ↓ ｽﾌﾟﾚｯﾄﾞｼｰﾄの通知済みか否かをfalse→trueにする
      // ------------------------------
      sh.getRange(2 + i, 8, 1, 1).setValue(true); // iはｲﾝﾃﾞｯｸｽ（行番号-1）なのでカラム名も含め+2が必要 8は2次の通知済みか否かの列番号



      break;
    }else{
      console.log("このイベントは今回の通知対象でない");
    }
  };
  Logger.log(message);
  return message;
}


/**
 * LINENotifyAPIで実際にメッセージを送信する
 * @param [string] message - LINEに送信するメッセージ
 * @returns [bool] - 送信に成功したか否か
 */
function sendNotify(message = "メッセージ送信テストです"){
  try{
    var url = 'https://notify-api.line.me/api/notify';
    var options = {
      'method' : 'post',
      'headers' : {
        'Authorization' : 'Bearer ' + PropertiesService.getScriptProperties().getProperty("LINE_NOTIFY_TOKEN_GROUP2") // カレンダー通知
      },
      'payload' : {
        'message' : message // createMessage()で作成した文字列
      }
    };
    UrlFetchApp.fetch(url, options);
  }catch(error){
    return false;
  }
  return true;
}