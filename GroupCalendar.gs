// シンプル組織カレンダー
function groupCalendar() {
  // カレンダーリスト
  var cals = [
    "example.jp_XXXXX@group.calendar.google.com", // 全体行事
    "boss@example.jp",
    "manager@example.jp",
    "staff@example.jp",
    "intern@example.jp"
    ];
  // スプレッドシート
  var spreadsheet = SpreadsheetApp.openByUrl("https://docs.google.com/spreadsheets/d/XXXXX/edit#gid=0");
  var sheets = spreadsheet.getSheets();
  var sheet = sheets[0];
  var weekDayList = [ "日", "月", "火", "水", "木", "金", "土" ] ;
  var color;
  // 日本の祝日
  var holidaysCal = CalendarApp.getCalendarById("ja.japanese#holiday@group.v.calendar.google.com");	
  
  for (var j = 0; j < cals.length; j++) {
    var cal = CalendarApp.getCalendarById( cals[j] );	// カレンダー切替
    for (var i = 0; i < 5 ; i++) {	　　// 対象は5日先まで
      // 列見出し（名称）
      if ( i == 0 ) {
        var name = cal.getName();
        Logger.log(name);
      }
      var target = new Date();
      target.setDate(target.getDate() + i);	// 取得日設定
      // 行見出し（月日曜）
      if ( j == 0 ) {
        var month = target.getMonth() + 1;	// ※月は0始まり
        var date = target.getDate();  // 日
        var weekDay = weekDayList[ target.getDay() ];	// 曜日
        var holidayEvents = holidaysCal.getEventsForDay( target );
        // 曜日に色付け
        if (holidayEvents.length) {
          // 祝日
          color = "magenta";
        } else if (weekDay == "日") {
          color = "red";
        } else if (weekDay == "土") {
          color = "blue";
        } else {
          color = "black";
        }
        sheet.getRange(1 + j, 2 + i).setFontColor(color);
        var gappi = month + "月" + date + "日（" + weekDay + "）" ;
        if (holidayEvents.length) {
          for each (var e in holidayEvents) {
            gappi = gappi + "\n" + e.getTitle();
          }
        }
        sheet.getRange(1 + j, 2 + i).setValue(gappi);
      }
      var events = cal.getEventsForDay(target);	// イベントを取得
      var title　= "";				// 初期化しないとundefinedとなるための対策
      for each(var evt in events) {
        if (title == "") {
          title = evt.getTitle();
        } else {
          title = title + "\n" + evt.getTitle();
        }
      }
      sheet.getRange(2 + j, 2 + i).setValue(title);	// タイトル
    }
  }
}
