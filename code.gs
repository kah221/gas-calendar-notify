/**
 * トリガーにより自動で実行される関数
 * 処理の流れ
 * - LINEで送るメッセージを作成
 * - LINENotifyAPIでメッセージを送信
 */
function executeWithTrigger(){
    // 送信するメッセージを作る
    message = createMessage();
    
    // LINE Notify でメッセージを送信する
    is_successed = sendNotify(message);
    console.log(is_successed);
  
  }
  
  /**
   * Googleカレンダーで更新が起きたときに実行される
   * GoogleカレンダーのカレンダーIDに対応する予定データを取得する関数
   * 
   */
  function getEventsFromGCalendar() {
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
    // const calendar_obj = CalendarApp.getCalendarById(PropertiesService.getScriptProperties().getProperty("GOOGLE_CALENDAR_ID_R"));
    const calendar_obj = CalendarApp.getCalendarById(PropertiesService.getScriptProperties().getProperty("GOOGLE_CALENDAR_ID_M"));
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
      events_obj.forEach(function(event)  {
        // ------------------------------
        // ↓ イベントの諸情報を変数として取得
        // ------------------------------
        var title = event.getTitle();         // イベントのタイトルを取得
        var start = event.getStartTime();     // イベントの開始時刻を取得 Dateｵﾌﾞｼﾞｪｸﾄ
        var end = event.getEndTime();         // イベントの終了時刻を取得
        var isAllday = event.isAllDayEvent(); // イベントが終日イベントかどうかを取得
        var isNoticed = false;                // 既に通知済みか否か
        var trigger_time = '';                // ｽﾌﾟﾚｯﾄﾞｼｰﾄに記述するため空文字で用意（終日イベ）
  
        // ------------------------------
        // トリガー
        // ------------------------------
        if(!isAllday){ // 終日予定でない場合のみトリガーを追加
          var now_time = new Date();  // 現在時刻を取得
          if(start < now_time){       // この関数が実行されたときにイベントの開始時刻が既に過去の場合 → トリガーを作らない、既に通知済みか否かをfalseにする
            isNoticed = true;
          }else{ // まだイベント開始時刻を迎えていない場合←★
            // ｽﾌﾟﾚｯﾄﾞｼｰﾄから何分前に通知させるか「notice」を取得
            var notice = Number(ss.getSheetByName("設定").getRange("B1").getValue()); // 必ず半角数値5~60までの間であるとする
            
            // トリガーの時刻を作成
            var trigger_time = new Date(start);
            trigger_time.setMinutes(trigger_time.getMinutes() - notice); // トリガーの時刻を開始時刻「start」からnotice分だけ早める
  
            // ★且つ、現在が「予定開始と通知時刻の間」のとき、トリガーは作成せず、即時通知を行う
            if(trigger_time < now_time){
              console.log("即時通知対象イベント");
              isExistInstanceEvent = true; // フラグを立てる
                // ここで、G列に対応するisNoticedはfalseのままにしておく
            }else{
              // 次回の実行時間を指定（トリガー機能）
              ScriptApp.newTrigger('executeWithTrigger').timeBased().at(trigger_time).create();
            }
          }
        }else{ // 既にイベント開始時刻を迎えていた場合
          isNoticed = true; // 終日予定の場合予めtrueにしておき通知を行わない。トリガーも作成しない
        }
  
        // ------------------------------
        // ↓ 二次元配列に格納
        // ------------------------------
        /**
         * 予定のタイトル [string]
         * 予定の開始時刻 [Dateｵﾌﾞｼﾞｪｸﾄ]
         * 予定の終了時刻 [Dateｵﾌﾞｼﾞｪｸﾄ]
         * 終日予定か否か [bool] trueなら終日予定
         * カレンダー識別用文字列 [string] ← 結局使っていない
         * ポインター [bool] false←まだい通知を行っていない。true←すでに通知を行った又は終日予定
         * ★イベントの長さはメッセージ作成時に計算を行うのでここでは書かない
         * */ 
        events_arr.push([title, start, end, trigger_time, isAllday, '01', isNoticed]);
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
    // ↑ 書込を行ってから、
    // ↓ 即時通知対象イベントがある場合、通知を行う
    //   （ｽﾌﾟﾚｯﾄﾞｼｰﾄG列のfalseをsendNotifyでtrueに書き換える）
    // ------------------------------
    if(isExistInstanceEvent == true){
      message = createMessage(true); // 即時通知イベントなので引数にtrueを指定する
      result = sendNotify(message);
      // console.log(result);
    }
  
    return;
  }
  
  
  
  
  /**
   * LINEに送信するメッセージ文字列を作成する関数
   * 処理の流れ
   * - ｽﾌﾟﾚｯﾄﾞｼｰﾄからイベントをすべて取得
   * - ↑をループで1行ずつ読み取る
   * - 「G列：通知済みか否か」がfalseのものについてデータを変数に格納しループを出る
   * - 変数と、ｽﾌﾟﾚｯﾄﾞｼｰﾄに定義したフォーマットに従いにメッセージを作成
   * - ｽﾌﾟﾚｯﾄﾞｼｰﾄの「G列：通知済みか否か」をtrueに上書き
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
    var notice = Number(ss.getSheetByName("設定").getRange("B1").getValue());              // 必ず半角数値5~60までの間であるとする
    var events = sh.getRange(2, 1, sh.getLastRow() - 1, sh.getLastColumn()).getValues();  // イベントを2次元配列で取得 相対座標なのでrowは-1
    Logger.log(events);
  
    // ------------------------------
    // ↓ ループで1行ずつ読み取る
    // ------------------------------
    for(var i=0; i<events.length; i++){
      event = events[i];
      if(event[4] == false && event[6] == false){ // 終日予定でない且つ通知済みか否かがfalseのイベントが初めて出てきたとき、そのイベントが今回の通知対対象
        Logger.log('通知対象');
  
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
        event_stt = st_h + ':' + st_m; // 開始時刻
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
          message += "【" + notice + "分後】\n\n"; // 9
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
        sh.getRange(2 + i, 7, 1, 1).setValue(true); // iはｲﾝﾃﾞｯｸｽ（行番号-1）なのでカラム名も含め+2が必要 7は通知済みか否かの列番号
  
  
        break; // 通知済みか否かの部分がfalseの部分が来た時に抜け出す
      }else{
        Logger.log('通知対象でない');
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
    // 最大文字数：トークンの名前で全角7文字を使った場合、残り半角40文字
    try{
      var url = 'https://notify-api.line.me/api/notify';
      var options = {
        'method' : 'post',
        'headers' : {
          // 'Authorization' : 'Bearer ' + PropertiesService.getScriptProperties().getProperty("LINE_NOTIFY_TOKEN")
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