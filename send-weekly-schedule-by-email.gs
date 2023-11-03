/*
参考情報

- Googleカレンダーの1週間分の予定をGASを使いLINE通知
https://qiita.com/y-araki-qiita/items/26224c9ca2628d6e4d65

- GASでメール送信する際の各種オプション
https://blog.synnex.co.jp/google/options-sending-email-by-gas/
*/

/*
注意事項

複数カレンダーの予定を送る場合は、カレンダーの予定をあらかじめ共有しておくこと。
例えば、カレンダーA, B, Cが存在するとする。
この時、カレンダーA経由でBとCの予定も通知する場合は、BとCの予定をAに共有しておく必要がある。
*/

function myFunction(){

  var startDate = new Date();
  var period = 7; // 予定を表示したい日数

  startDate.setDate(startDate.getDate());
  var schedule = getSchedule(startDate, period);
  var message = scheduleToMessage(schedule, startDate, period);

  sendScheduleByEmail(message);
}

// Googleカレンダーから予定取得
function getSchedule(startDate, period){
  var calendarIDs = ['']; //カレンダーIDの配列。カレンダーを共有している場合、複数指定可。

  var schedule = new Array(calendarIDs.length);
  for(var i=0; i<schedule.length; i++){schedule[i] = new Array(period);}

  for(var iCalendar=0; iCalendar < calendarIDs.length; iCalendar++){
    var calendar = CalendarApp.getCalendarById(calendarIDs[iCalendar]);

    var date = new Date(startDate);
    for(var iDate=0; iDate < period; iDate++){
      schedule[iCalendar][iDate] = getDayEvents(calendar, date);
      date.setDate(date.getDate() + 1);
    }
  }

  return schedule;
}  

function getDayEvents(calendar, date){
  var dayEvents = "";
  var events = calendar.getEventsForDay(date);

  for(var iEvent = 0; iEvent < events.length; iEvent++){
    var event = events[iEvent];
    var title = event.getTitle();
    var startTime = _HHmm(event.getStartTime());
    var endTime = _HHmm(event.getEndTime());

    dayEvents = dayEvents + '・' + startTime + '-' + endTime + ' ' + title + '\n';
  }

  return dayEvents;
}

function scheduleToMessage(schedule, startDate, period){
  var now = new Date();
  var body = '\n1週間の予定\n' 
    + '(' +_Md(now) + ' ' +_HHmm(now) + '時点)\n'
    + '--------------------\n';

  var date = new Date(startDate);
  for(var iDay=0; iDay < period; iDay++){
    body = body + _Md(date) + '(' + _JPdayOfWeek(date) + ')\n';

    for(var iCalendar=0; iCalendar < schedule.length; iCalendar++){
      body = body + schedule[iCalendar][iDay];
    }   

    date.setDate(date.getDate() + 1);
    body = body + '\n';
  }

  return body;
}

// メール送信
function sendScheduleByEmail(message) {
  const address = '送信先メールアドレス';

  const subject = '1週間の予定通知';

  const options = { 
    from: address,
    name: 'プライベートスケジュール配信'
   };

  let body = message;

  GmailApp.sendEmail(address, subject, body, options);
}

// 日付指定の関数
function _HHmm(date){
  return Utilities.formatDate(date, 'JST', 'HH:mm');
}

function _Md(date){
  return Utilities.formatDate(date, 'JST', 'M/d');
}

function _JPdayOfWeek(date){
  var dayStr = ['日', '月', '火', '水', '木', '金', '土'];
  var dayOfWeek = date.getDay();

  return dayStr[dayOfWeek];
}
