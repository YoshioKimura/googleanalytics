function getGaDataPrevDay() {

//-------日時データの取得-----------
// 昨日の日付を取得する
var date = new Date();
date.setDate(date.getDate() - 1);

//年度を取得する
var year = date.getFullYear();

//曜日を取得する
var ary = ['Sun', 'Mon', 'Tue', 'Wed', 'Thr', 'Fri', 'Sat'];
var weekDayNum = date.getDay();//曜日番号の取得
var weekDay = ary[weekDayNum] ;

//-------GAデータ-----------
// GAからデータを取得する関数の定義
function getGaData(startDate, endDate) {
var gaData = Analytics.Data.Ga.get(
'ga:', //アナリティクスで使っているID名を記載
startDate,
endDate,
'ga:pageviews, ga:sessions, ga:users, ga:newUsers,ga:bounceRate,ga:exitRate,ga:avgSessionDuration,ga:pageviewsPerSession, ga:totalEvents'
).rows;

return gaData;
}

// 戻り値（指定期間）を指定し、関数を実効
var gaData = getGaData('yesterday', 'yesterday');

//-------スプレッドシートへの書き込み-----------
// 書き込むシートを選択する
var spreadSheet = SpreadsheetApp.getActiveSpreadsheet();
var sheet = spreadSheet.getSheets()[0]; 

//最終行+1行目の取得
var lastRow = sheet.getLastRow()+1;
//最終行+1行目を追加
sheet.insertRows(lastRow);

//追加した行に書き込む
sheet.getRange(lastRow, 1).setValue(year);
sheet.getRange(lastRow, 2).setValue(date);
sheet.getRange(lastRow, 3).setValue(weekDay);
sheet.getRange(lastRow, 4).setValue(gaData[0][0]);
sheet.getRange(lastRow, 5).setValue(gaData[0][1]);
sheet.getRange(lastRow, 6).setValue(gaData[0][2]);
sheet.getRange(lastRow, 7).setValue(gaData[0][3]);
sheet.getRange(lastRow, 8).setValue(gaData[0][4] + "%\n");
sheet.getRange(lastRow, 9).setValue(gaData[0][5] + "%\n");
sheet.getRange(lastRow, 10).setValue(gaData[0][6]/60);
sheet.getRange(lastRow, 11).setValue(gaData[0][7]);
sheet.getRange(lastRow, 12).setValue(gaData[0][8]);

//-----セルのフォーマット指定-------
//日時データのフォーマット
var FormatDate = sheet.getRange("B:B").setNumberFormat("MM/DD");
//直帰率
//var FormatBounceRate = sheet.getRange("H:H").setNumberFormat("0.00%");
//滞在時間
var FormatDuration = sheet.getRange("J:J").setNumberFormat("#,##00.00");
//ページ/セッション
var FormatPageSession = sheet.getRange("K:K").setNumberFormat("#,##0.000");
}
