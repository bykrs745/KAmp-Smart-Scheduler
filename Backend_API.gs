/**
 * Backend_API.gs
 * フロントエンドから `google.script.run` 経由で呼び出されるAPI群
 */

function getDayInfoForWeb(dateStr) {
  const parsedDate = new Date(dateStr.replace(/-/g, "/"));
  if (isNaN(parsedDate.getTime())) return { error: "Invalid Date" };
  
  let holiday = "";
  let seasonal = "";
  
  try {
    const holidays = CalendarApp.getCalendarById('ja.japanese#holiday@group.v.calendar.google.com').getEventsForDay(parsedDate);
    if (holidays.length > 0) {
      const title = holidays[0].getTitle();
      const nationalHolidays = ["元日", "成人の日", "建国記念の日", "天皇誕生日", "春分の日", "昭和の日", "憲法記念日", "みどりの日", "こどもの日", "海の日", "山の日", "敬老の日", "秋分の日", "スポーツの日", "文化の日", "勤労感謝の日", "振替休日", "国民の休日"];
      
      let isNationalHoliday = false;
      for (let i = 0; i < nationalHolidays.length; i++) {
        if (title.indexOf(nationalHolidays[i]) !== -1) {
          isNationalHoliday = true;
          break;
        }
      }
      if (isNationalHoliday) {
        holiday = title;
      } else {
        seasonal = title;
      }
    }
  } catch(e) {}
  
  const day = parsedDate.getDay();
  let dayType = "平日";
  let colorClass = "text-white";
  if (holiday !== "" || day === 0) {
    dayType = "日祝日";
    colorClass = "text-red-400";
  } else if (day === 6) {
    dayType = "土曜日";
    colorClass = "text-blue-400";
  }

  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sheet = ss.getSheetByName("時刻表マスタ");
  const busTimes = [];
  if (sheet) {
    const data = sheet.getDataRange().getValues();
    for (let i = 1; i < data.length; i++) {
      const cat = data[i][1] ? data[i][1].toString().trim() : "";
      let masterTime = data[i][4] ? data[i][4].toString().replace(/'/g, "").trim() : "";
      let mParts = masterTime.split(":");
      if (mParts.length >= 2) {
        masterTime = ("0" + parseInt(mParts[0], 10)).slice(-2) + ":" + ("0" + parseInt(mParts[1], 10)).slice(-2);
      }
      if (cat.indexOf(dayType) !== -1 || cat === "") {
        if(masterTime) busTimes.push(masterTime);
      }
    }
  }
  if (busTimes.length === 0) {
    busTimes.push("08:00", "08:30", "09:00", "09:30", "10:00");
  }
  const uniqueBusTimes = busTimes.filter((x, i, self) => self.indexOf(x) === i).sort();
  
  return { holiday: holiday, seasonal: seasonal, dayType: dayType, colorClass: colorClass, busTimes: uniqueBusTimes };
}

function getRoutesForWeb(dateStr, timeStr, destinationAddr) {
  return calculateTransitRouteDetailsMultiple(timeStr, dateStr, destinationAddr);
}

function getLatestPlaceDetails(placeId, name, address) {
  const settings = getSettings();
  let latestInfo = null;
  if (settings.API_KEY && placeId) {
    const results = searchMasterOrApi(name + " " + address, settings.API_KEY);
    for(let i=0; i<results.length; i++){
      if(results[i].placeId === placeId) {
        latestInfo = results[i];
        break;
      }
    }
    if(!latestInfo && results.length > 0) latestInfo = results[0];
  }
  return { success: !!latestInfo, data: latestInfo };
}

function getMasterListForWeb() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const masterSheet = ss.getSheetByName("目的地一覧マスタ");
  if (!masterSheet) return [];
  const data = masterSheet.getDataRange().getValues();
  const list = [];
  for (let i = 1; i < data.length; i++) {
    if (data[i][1]) {
      list.push({
        placeId: data[i][0] ? data[i][0].toString().trim() : "",
        name: data[i][1] ? data[i][1].toString().trim() : "",
        address: data[i][2] ? data[i][2].toString().split('\n')[0] : "",
        stayTime: data[i][3] ? data[i][3].toString().trim() : "",
        calType: data[i][4] ? data[i][4].toString().trim() : "",
        startTime: data[i][5] ? data[i][5].toString().trim() : "",
        busTime: data[i][6] ? data[i][6].toString().trim() : ""
      });
    }
  }
  return list;
}

/**
 * 🌟 新規追加: Homeシートから登録済みの予定一覧（Deck用データ）を取得する
 */
function getHomeListForWeb() {
  try {
    const ss = SpreadsheetApp.getActiveSpreadsheet();
    const sheet = ss.getSheetByName("Home");
    if (!sheet) return [];
    
    const colMap = getColumnMap(sheet);
    const lastRow = sheet.getLastRow();
    if (lastRow < 2) return [];
    
    const list = [];
    // 直近の登録を上に表示するため、下から逆順にループします
    for (let r = lastRow; r >= 2; r--) {
      const rowObj = getRowDataAsObject(sheet, r, colMap);
      if (rowObj["名前"]) {
        list.push({
          id: rowObj["ユニークID"] || ("item_" + r),
          name: rowObj["名前"] ? rowObj["名前"].toString().trim() : "名称未設定",
          address: rowObj["住所"] ? rowObj["住所"].toString().trim() : "",
          date: rowObj["出発日"] ? Utilities.formatDate(new Date(rowObj["出発日"]), Session.getScriptTimeZone(), "yyyy/MM/dd") : "日付未設定",
          busTime: rowObj["バス停時間"] ? rowObj["バス停時間"].toString().replace(/'/g, "").trim() : "",
          stayTime: rowObj["滞在時間"] ? rowObj["滞在時間"].toString().trim() : "",
          arrTime: rowObj["目的地到着時間"] ? rowObj["目的地到着時間"].toString().replace(/'/g, "").trim() : "",
          retTime: rowObj["帰宅時間"] ? rowObj["帰宅時間"].toString().replace(/'/g, "").trim() : "",
          calType: rowObj["カレンダー種別"] ? rowObj["カレンダー種別"].toString().trim() : "",
          eventId: rowObj["イベントID"] ? rowObj["イベントID"].toString().trim() : ""
        });
      }
    }
    return list;
  } catch (e) {
    Logger.log("getHomeListForWeb Error: " + e.message);
    return [];
  }
}

function executeWebSearch(keyword) {
  const settings = getSettings();
  let results = [];
  if (settings.API_KEY) {
    results = searchMasterOrApi(keyword, settings.API_KEY);
  }
  
  if (results.length === 1) {
    return { success: true, message: "1件見つかり、自動選択されました", data: results };
  } else if (results.length > 1) {
    return { success: true, message: "候補が " + results.length + " 件見つかりました。選択してください。", data: results };
  } else {
    return { success: false, message: "該当がありませんでした", data: [] };
  }
}

function executeFullAutomation(data, rowNumber) {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sheet = ss.getSheetByName("Home");
  const colMap = getColumnMap(sheet);
  
  // 新規登録時は最終行の下に追加
  const targetRow = sheet.getLastRow() + 1;
  
  const valText = data.name + " / " + data.address + " / " + data.placeId;
  const uniqueId = "ID_" + new Date().getTime() + "_" + Math.floor(Math.random() * 1000);
  
  const valuesToSet = {
    "ユニークID": uniqueId,
    "選択": valText,
    "名前": data.name,
    "郵便": data.postal,
    "住所": data.address,
    "建物": data.building,
    "電話": data.phone,
    "WEB": data.website,
    "営業時間": data.bizHours,
    "PlaceID": data.placeId,
    "出発日": data.date,
    "滞在時間": data.stayTime,
    "祝日": data.holiday,
    "季節の行事": data.seasonalEvent,
    "出発時間": data.startTime,
    "バス停時間": data.busTime,
    "目的地到着時間": data.arrTime,
    "帰宅時間": data.retTime,
    "カレンダー種別": data.calType,
    "天気1": "🌤️", // デフォルト天気プレースホルダー
    "INFO": "WEBアプリより登録完了"
  };
  
  for (let colName in valuesToSet) {
    let colIndex = colMap[colName];
    if (colIndex) {
      sheet.getRange(targetRow, colIndex).setValue(valuesToSet[colName]);
    }
  }
  
  SpreadsheetApp.flush();
  return { success: true, message: "スプレッドシートへの予定登録が完了しました！" };
}

function getWeatherForWeb(address, dateStr, arrivalTimeVal) {
  const settings = getSettings();
  const parsedDate = new Date(dateStr.replace(/-/g, "/"));
  return getWeatherData(address, parsedDate, arrivalTimeVal, settings.OPENWEATHER_API_KEY);
}