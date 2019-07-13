//======================================================
// 設定値
// 設置値が記載されたシート名
var SETTING_SHEET_NAME = "setting";

// 自己都合休日の日付の先頭セル
var PERSONAL_HOLIDAY_START_CELL = {
    row: 3,
    column: 2
};
// 時間情報の先頭セル
var TASK_HOUR_START_CELL = {
    row: 3,
    column: 2
};
// 開始日付情報の先頭セル
var TASK_DATE_START_CELL = {
    row: 3,
    column: 3
};

// 終了日付情報の先頭セル
var TASK_DATE_END_CELL = {
    row: 3,
    column: 4
};

// 一日の仕事時間
var TASK_HOUR_BY_DAY = {
    row: 3,
    column: 3
};
//====================================================

// 現在のスプレッドシートを取得
var spreadsheet = SpreadsheetApp.getActiveSpreadsheet();

// WBS シートを取得
var wbsSheet = spreadsheet.getActiveSheet();

// 設定値 シートを取得
var settingSheet = spreadsheet.getSheetByName(SETTING_SHEET_NAME);
// 自己都合休日を取得
var holidays = settingSheet.getRange(
    PERSONAL_HOLIDAY_START_CELL.row,
    PERSONAL_HOLIDAY_START_CELL.column,
    settingSheet.getLastRow()).getValues();

// Calender から祝日情報を取得
var calendars = CalendarApp.getCalendarsByName('日本の祝日');


function main() {

    // 1日の仕事時間
    const HOURS_BY_DAY = settingSheet.getRange(TASK_HOUR_BY_DAY.row, TASK_HOUR_BY_DAY.column).getValue();
    // タスク開始の日付
    const START_DATE = wbsSheet.getRange(TASK_DATE_START_CELL.row, TASK_DATE_START_CELL.column).getValue();
    // タスクの時間情報を取得
    const taskHours = wbsSheet.getRange(
        TASK_HOUR_START_CELL.row,
        TASK_HOUR_START_CELL.column,
        wbsSheet.getLastRow()).getValues();
    // 時間が格納されているセルの数
    var taskSize = taskHours.filter(String).length;
    // 現在の日付
    var currentDay = new Date(START_DATE);
    // 現在の日付の累積仕事予定時間
    var accumulatedHourOfDay = 0;
    // 現在のセル情報
    var currentCell = wbsSheet.getRange(TASK_DATE_START_CELL.row, TASK_DATE_START_CELL.column);
    // 開始日を次の日に設定するフラグ
    var incrementStartDay = false;


    for (var i = 0; i < taskSize; i++) {
        // 仕事時間
        var taskHour = taskHours[i][0];

        // タスクの時間が存在しない場合には終了
        if (taskHour.length == 0) break;

        if (incrementStartDay == true) {
            // 次の日を取得
            currentDay = getBusinessDay(currentDay, 1);
        }

        // 開始日を設定
        currentCell.setValue(currentDay);
        // 終了日 - 開始日の日数
        var offsetDay = Math.floor((accumulatedHourOfDay + taskHour) / HOURS_BY_DAY);
        // 終了日の累積労働時間を算出
        accumulatedHourOfDay = (accumulatedHourOfDay + taskHour) - offsetDay * HOURS_BY_DAY;

        if (offsetDay > 0 && accumulatedHourOfDay == 0) {
            // 開始日を次の日に設定するか否か
            incrementStartDay = true;
            offsetDay -= 1;
        } else {
            incrementStartDay = false;
        }

        // 終了日の日付を取得
        currentDay = getBusinessDay(currentDay, offsetDay);

        // 終了日を設定
        currentCell.offset(0, TASK_DATE_END_CELL.column - currentCell.getColumn()).setValue(currentDay);
        // 現在のセルを一つ下の開始日のセルに設定
        currentCell = currentCell.offset(1, 0);
    }
}

/**
 * 指定した日の次の営業日を取得する。
 */
function getBusinessDay(currentDay, offsetDay) {

    var bussinessDay = currentDay;
    // 祝日か否か
    var isPublicHoliday = false;
    // 指定した日付を返す
    if (offsetDay == 0) return bussinessDay;

    for (var i = 0; i < offsetDay; i++) {
        do {
            // 次の日を取得
            bussinessDay = new Date(
                bussinessDay.getYear(),
                bussinessDay.getMonth(),
                bussinessDay.getDate() + 1
            );

            // 曜日を取得
            var day = bussinessDay.getDay();
            // 祝日か否か
            isPublicHoliday = (calendars[0].getEventsForDay(bussinessDay).length > 0);
            // 休日の場合には再取得
        } while (day == 0 || day == 6 || isHoliday(bussinessDay) || isPublicHoliday);

    }

    // 営業日
    return bussinessDay;
}

/**
 * 指定した日が自己都合休日日かをを取得
 * @param {*} currentDay 
 */
function isHoliday(currentDay) {

    // 休日が設定されていない場合には終了
    if (holidays.length == 0) return false;

    // 休日か調べたい日付を UnixTimeStamp へ変換
    var currentDayTime = Math.round(currentDay.getTime() / 86400000);

    for (var i = 0; i < holidays.length; i++) {
        // 休日
        var holiday = holidays[i][0];
        // Date 型でない場合は除外
        if (Object.prototype.toString.call(holiday).slice(8, -1) !== "Date") continue;

        // 休日の日付を UnixTimeStamp へ変換
        var holidayTime = Math.round(holiday.getTime() / 86400000);

        // 一致した場合は 休日
        if (holidayTime === currentDayTime) return true;
    }
    // 休日でない
    return false;
}