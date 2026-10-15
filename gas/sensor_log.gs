const SENSOR_SHEET_NAME = 'センサーログ';

// 列ヘッダー（グループ別にまとめた並び）
const HEADERS = [
  '日時',
  // pH
  'pH 棚1A', 'pH 棚1B', 'pH 棚2A', 'pH 棚2B',
  // EC
  'EC 棚1A', 'EC 棚1B', 'EC 棚2A', 'EC 棚2B',
  // 水温
  '水温 棚1A', '水温 棚1B',
  // 気温
  '気温 おんどとり1', '気温 おんどとり2', '気温 air1', '気温 air2',
  // 湿度
  '湿度 おんどとり1', '湿度 おんどとり2', '湿度 air1', '湿度 air2',
  // CO2
  'CO2 おんどとり1', 'CO2 おんどとり2', 'CO2 air1', 'CO2 air2',
];

function doPost(e) {
  try {
    const json = JSON.parse(e.postData.contents);
    const ss = SpreadsheetApp.openById(SPREADSHEET_ID);
    const sheet = ss.getSheetByName(SENSOR_SHEET_NAME);

    // ヘッダーがなければ書く
    if (sheet.getLastRow() === 0) {
      sheet.appendRow(HEADERS);
    }

    const now = new Date();
    const jst = new Date(now.getTime() + 9 * 60 * 60 * 1000);
    const ts = Utilities.formatDate(jst, 'Asia/Tokyo', 'yyyy/MM/dd HH:mm');

    // センサー値をパース（値は /100 or /10 で実値に変換）
    const p4 = json.ECpH4   || {};  // 棚1A pH
    const p2 = json.ECpH2   || {};  // 棚1B pH
    const p3 = json.ECpH3   || {};  // 棚1A EC
    const w1 = json.waterSensors1 || {};  // 棚2A
    const w2 = json.waterSensors2 || {};  // 棚2B
    const t1 = json.tr76ul1  || {};  // おんどとり1
    const t2 = json.tr76ul2  || {};  // おんどとり2
    const a1 = json.airSensor1 || {};
    const a2 = json.airSensor2 || {};

    const row = [
      ts,
      // pH
      p4.ph   != null ? p4.ph / 100   : '',
      p2.ph   != null ? p2.ph / 100   : '',
      w1.pH   != null ? w1.pH          : '',
      w2.pH   != null ? w2.pH          : '',
      // EC
      p3.ec   != null ? p3.ec / 100   : '',
      '',  // 棚1B EC（センサーなし）
      w1.EC   != null ? w1.EC          : '',
      w2.EC   != null ? w2.EC          : '',
      // 水温
      p4.temp != null ? p4.temp / 10  : '',
      p2.temp != null ? p2.temp / 10  : '',
      // 気温
      t1.temp != null ? t1.temp        : '',
      t2.temp != null ? t2.temp        : '',
      a1.temp != null ? a1.temp        : '',
      a2.temp != null ? a2.temp        : '',
      // 湿度
      t1.humid != null ? t1.humid      : '',
      t2.humid != null ? t2.humid      : '',
      a1.humid != null ? a1.humid      : '',
      a2.humid != null ? a2.humid      : '',
      // CO2
      t1.co2  != null ? t1.co2         : '',
      t2.co2  != null ? t2.co2         : '',
      a1.co2  != null ? a1.co2         : '',
      a2.co2  != null ? a2.co2         : '',
    ];

    sheet.appendRow(row);

    return ContentService
      .createTextOutput(JSON.stringify({ status: 'ok', ts }))
      .setMimeType(ContentService.MimeType.JSON);

  } catch (err) {
    return ContentService
      .createTextOutput(JSON.stringify({ status: 'error', message: err.message }))
      .setMimeType(ContentService.MimeType.JSON);
  }
}
