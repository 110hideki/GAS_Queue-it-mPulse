function myFunction() {
  var spreadsheet = SpreadsheetApp.getActiveSpreadsheet();
  // Q_mPulse という名称の Tab を事前に作成する
  var sheet = spreadsheet.getSheetByName('Q_mPulse');
  if (!sheet) return;

  // 入力列の設定(行を挿入して設定)
  sheet.insertRowBefore(3);
  var dateCell = sheet.getRange(3, 1);
  // 10/25/2022 2:13:00
  dateCell.setValue(
    Utilities.formatDate(new Date(), 'Asia/Tokyo', 'yyyy/MM/dd HH:mm')
  );

  var d_d1 = new Date();

    // var qitformatdate = Utilities.formatDate(new Date(), 'GMT', "yyyy-MM-dd'T'HH:MM':00Z'")
    // 分単位で一分前までのデータが取れるので、現在の時間から一分前の一分間のデータを入手する
    // QIT from=2022-10-24T07:00:00Z&to=2022-10-24T07:01:00Z
    // yyyy-MM-ddThh:mm:ssZ
    var date_E = new Date(d_d1.getFullYear(),d_d1.getMonth(),d_d1.getDate(),d_d1.getHours(),d_d1.getMinutes() - 1,0);
    var date_S = new Date(d_d1.getFullYear(),d_d1.getMonth(),d_d1.getDate(),d_d1.getHours(),d_d1.getMinutes() - 2,0);
    var ft_endtime = Utilities.formatDate(date_E, 'GMT', "yyyy-MM-dd'T'HH:mm':00Z'")
    var ft_starttime = Utilities.formatDate(date_S, 'GMT', "yyyy-MM-dd'T'HH:mm':00Z'")

//  console.log(Utilities.formatDate(d_d1, 'Asia/Tokyo', 'yyyy/MM/dd HH:mm'));

//２次元配列として設定
  var arr = [[]];
  var apiResult = arr;
  var rt_q_data;
  var rt_ttfb_data;
 
  // Queue-it Data
  if ((rt_q_data = get_json_qit_qstat("maxoutflow", ft_starttime, ft_endtime)) === undefined) return;
  qit_convertArray(arr, rt_q_data);

  if ((rt_q_data = get_json_qit_qstat("queueuniqueinflow", ft_starttime, ft_endtime)) === undefined) return;
  qit_convertArray(arr, rt_q_data);

  if ((rt_q_data = get_json_qit_qstat("queueidsinqueue", ft_starttime, ft_endtime)) === undefined) return;
  qit_convertArray(arr, rt_q_data);

  if ((rt_q_data = get_json_qit_qstat("queueexpectedwaittime", ft_starttime, ft_endtime)) === undefined) return;
  qit_convertArray(arr, rt_q_data);

  if ((rt_q_data = get_json_qit_qstat("queueuniqueoutflow", ft_starttime, ft_endtime)) === undefined) return;
  qit_convertArray(arr, rt_q_data);

  if ((rt_q_data = get_json_qit_qstat("safetynetoutflow", ft_starttime, ft_endtime)) === undefined) return;
  qit_convertArray(arr, rt_q_data);

  if ((rt_q_data = get_json_qit_qstat("queueoutflow", ft_starttime, ft_endtime)) === undefined) return;
  qit_convertArray(arr, rt_q_data);

  // mPulse Data
  // TTFB 75 Percentile の取得
  if ((rt_ttfb_data = get_json_mpulse_ttfb(ft_starttime, ft_endtime, 75)) === undefined) return;
  mpulse_convertArray(arr, rt_ttfb_data);

  // TTFB 50 Percentile の取得
  if ((rt_ttfb_data = get_json_mpulse_ttfb(ft_starttime, ft_endtime, 50)) === undefined) return;
  mpulse_convertArray(arr, rt_ttfb_data);


  //結果の貼り付け
  sheet
   .getRange(3, 2, apiResult.length, apiResult[0].length)
  .setValues(apiResult);
}

//
// Queue-it API を使って JSON を取得し、シートに出力するために必要な項目だけを抽出する
//
function get_json_qit_qstat(staticstype, starttime, endtime) {
  var request = qit_get_q_summary(staticstype, starttime, endtime);

  var options = {
  'method' : 'get',
   headers: {
   "api-key": PropertiesService.getScriptProperties().getProperty('queueit_APIKEY')
   },
   muteHttpExceptions: false,
   escaping: true
  };

   try {
      var response = UrlFetchApp.fetch(request, options);
   } catch (err) {
    Logger.log(err);
    return err;
  }
  
  var responseCode = response.getResponseCode();
  var responseBody = response.getContentText();

  //　Fetch に失敗することがあるので、そのときはログにデバッグ情報を記録する。View Logs でログが見える。
  if (responseCode !== 200) {
    Logger.log(Utilities.formatString("Request failed. response code: %d, Fetched URL: %s", responseCode));
    Logger.log(Utilities.formatString("Request body: %s", responseBody));
    return;
  }
  // jsonをパースする
  var parsedResult = JSON.parse(responseBody);
  return parsedResult;
}

//
// mPulse API を使って JSON を取得し、シートに出力するために必要な項目だけを抽出する
//
function get_json_mpulse_ttfb(starttime, endtime, percentile) {
  var request = mpulse_get_ttfb(starttime, endtime, percentile);

  var options = {
  'method' : 'get',
   headers: {
      "Authentication": PropertiesService.getScriptProperties().getProperty('mpulse_Authentication')
   },
   muteHttpExceptions: false,
   escaping: true
  };

  try {
      var response = UrlFetchApp.fetch(request, options);
   } catch (err) {
      Logger.log(err);
      return err;
  }
  
  var responseCode = response.getResponseCode();
  var responseBody = response.getContentText();

//　Fetch に失敗することがあるので、そのときはログにデバッグ情報を記録する。
// View Logs でログが見える。
  if (responseCode !== 200) {
    Logger.log(Utilities.formatString("Request failed. response code: %d, Fetched URL: %s", responseCode));
    Logger.log(Utilities.formatString("Request body: %s", responseBody));
    return;
  }
  
//  // jsonをパースする
  var parsedResult = JSON.parse(responseBody);
 
  return parsedResult;
}

function qit_convertArray(arr, content) {
   if (content.Entries[0].Sum === undefined) {
      arr[0].push("");
      Logger.log(Utilities.formatString("qit_convertArray: undefined"));
      return arr; 
  } else {
      arr[0].push(content.Entries[0].Sum );
  }
//  return arr;
  return;
}

function mpulse_convertArray(arr, content) {
 
   if (content.median === undefined) {
      arr[0].push("");
      Logger.log(Utilities.formatString("mpulse_convertArray: undefined"));
      return arr; 
  } else {
      arr[0].push(content.median);
  }
   return;
}

//
// Queue-it APIクエリーの作成
//
function qit_get_q_summary(staticstype, starttime, endtime) { 
  var api = 'https://' + PropertiesService.getScriptProperties().getProperty('queueit_ACCOUNT_DOMAIN') +
             '/2_0/event/tvc/queue/statistics/details/'; 
 // var fromto = "from=2022-10-24T07:00:00Z&to=2022-10-24T07:01:00Z";
  var fromto = 'from=' + starttime + '&' + 'to=' + endtime;
  var query = api + staticstype + '?' + fromto;
  return query;
}

//
// mPulse APIクエリーの作成
// https://mpulse.soasta.com/concerto/mpulse/api/v2/{mPulse_APIKEY}/summary?timezone=Asia/Tokyo&timer=FirstByte&percentile=75&date-start=2023-07-25T18:26:00Z&date-end=2023-07-25T18:27:00Z
//
function mpulse_get_ttfb(starttime, endtime, percentile) {
  var api = 'https://mpulse.soasta.com/concerto/mpulse/api/v2/' + 
                PropertiesService.getScriptProperties().getProperty('mpulse_APIKEY') +  
                '/summary?timezone=UTC&timer=FirstByte'
  var qs_fromto = 'date-start=' + starttime + '&' + 'date-end=' + endtime;
  var qs_percentile = 'percentile=' + percentile;
  var query = api + '&' + qs_percentile + '&' + qs_fromto;

  return query;
}

function doGet(e) { 
  myFunction();
  return ContentService.createTextOutput("Please refer to the spreadsheet");
}

