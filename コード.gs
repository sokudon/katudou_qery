
//共有が必要

//  1. Enter sheet name where data is to be written below           
        
//  2. Run > setup
//
//  3. Publish > Deploy as web app 
//    - enter Project Version name and click 'Save New Version' 
//    - set security level and enable service (most likely execute as 'me' and access 'anyone, even anonymously) 
//
//  4. Copy the 'Current web app URL' and post this in your form/script action 
//
//  5. Insert column names on your destination sheet matching the parameter names of the data you are passing in (exactly matching case)

var SCRIPT_PROP = PropertiesService.getScriptProperties(); // new property service

// If you don't want to expose either GET or POST methods you can comment out the appropriate function
function doPost(e) {
  // shortly after my original solution Google announced the LockService[1]
  // this prevents concurrent access overwritting data
  // [1] http://googleappsdeveloper.blogspot.co.uk/2011/10/concurrency-and-google-apps-script.html
  // we want a public lock, one that locks for all invocations
  var lock = LockService.getPublicLock();
  lock.waitLock(30000);  // wait 30 seconds before conceding defeat.

    var sname="カツドウSQL";//シート名
    
    
  try {
    // next set where we write the data - you could write to multiple/alternate destinations
    var doc = SpreadsheetApp.openById(SCRIPT_PROP.getProperty("key"));
    
     //JSONでーたを受け取り
     var JSON_DATA= JSON.parse(e.parameter["BD_JSON"]);
     var log="(*^_^*)でーた入力にご協力ありがとうございます(*^_^*)～\r\nごっごるしーと受け皿でのろぐをお伝えします～\r\n\r\n";        
     log= log+"送信したJSONでーた:\r\n" + JSON.stringify(JSON_DATA)+"\r\n\r\n";
    
    
    
    log= log+"対象しーと名: 時刻,すこあ,すてーたす,行番号(0以下は更新なし)\r\n";
    
     //しーと＋時刻べつに仕分けする    
    var sheet = doc.getSheetByName(sname);  
    var data =JSON_DATA.data;
    if(data.match(/QUERY\('ミリシタカツドウ'/)){
    sheet.getRange(1,1).setValue(data);
    
    var html =sddoGet();
    return HtmlService.createHtmlOutput(html);          
    }
    else{
    
    return "qery関数+ミリシタカツドウシート以外はつかえません";          
    }
    
  } catch(e){
    // 何か例外発生時のろぐ
    return ContentService
          .createTextOutput(JSON.stringify(JSON_DATA) + JSON.stringify({"結果":"エラー", "エラー": e}))
          .setMimeType(ContentService.MimeType.JSON);
  } finally { //release lock
    lock.releaseLock();
  }
}

function status(d){
var st;
if(d>0){
st=",成功,"
}
else if(d==0){
st=",同じすこあが存在,"
}
else if(d==-1){
st=",検索対象なし,"
}
else if(d==-2){
st=",送信すこあが前回より小さい,"
}

return st;
}

function setup() {
    var doc = SpreadsheetApp.getActiveSpreadsheet();
    SCRIPT_PROP.setProperty("key", doc.getId());
}

/**
 * Return a list of sheet names in the Spreadsheet with the given ID.
 * @param {String} a Spreadsheet ID.
 * @return {Array} A list of sheet names.
 */




function sddoGet() {
var sid="1CpwNLrurUVVLX2dmMgZHU-uQC7WQfyfWqLlaiooRaN8";
  
var sname="カツドウSQL";//シート名
  var ss = SpreadsheetApp.openById(sid);
  var sheets = ss.getSheetByName(sname);
  
　var last_row = 1;
　var last_col = 6;
  
  var valuea= sheets.getRange(1,1,600 ,1).getValues();
  for(var i=0;i<600;i++){
    if(valuea[i]==""){
    break;
    }
  }
  last_row=i;
  
  
  var values= sheets.getRange(1,1,last_row ,last_col).getValues();
 var value = JSON.parse(JSON.stringify(values));
 
  var html="";
  var moment=Moment.load();
  
  for(var i=0;i<last_row;i++){
    html+="<tr>";
  for(var k=0;k<last_col;k++){
    var tmp = value[i][k];
    if(k==0 && i >0){
      tmp= moment(tmp).format("YYYY/MM/DD HH:mm");
    }
   
    html+="<td>"+tmp+"</td>";
    
  }
  
    html+="</tr>";
  }
  
   
var sname="カツドウCAL";//シート名
  var ss = SpreadsheetApp.openById(sid);
  var sheets = ss.getSheetByName(sname);
  
　var last_row = 4;
　var last_col = 6;
  var html2="";
  
  var values= sheets.getRange(1,1,last_row ,last_col).getValues();
 var value = JSON.parse(JSON.stringify(values));
  
   for(var i=0;i<last_row;i++){
    html2+="<tr>";
  for(var k=0;k<last_col;k++){
    var tmp = value[i][k];
    if(k==0 && i >0 && i <3){
      tmp= moment(tmp).format("YYYY/MM/DD HH:mm");
    }
   
    html2+="<td>"+tmp+"</td>";
    
  }
  
    html2+="</tr>";
  }
 
  var header= "<style>th,td{  border:solid 1px #aaaaaa;},.table-scroll{  overflow-x : auto}</style>";
  
  
  html= "<table><tbody>"+ header+html + html2+ "<tbody></table>";
  
  
  return html;
}

//<>#"%　"#"はURI参照として、"%"はエスケープ用文字として使われます。
//除外されている記号 (RFC2396 に定義がないもの)
//以下の文字は使用できません。
// {}|\^[]`<>#"%

function hyperlink(link,a){
  link= "<a href='" + link +"' target=\"_blank\" rel=\"noopener\">" +a +"</a>";
  
  return link;
}

function wmap_getSheetsName(sheets){
  //var sheets = SpreadsheetApp.getActiveSpreadsheet().getSheets();
  var sheet_names = new Array();
  
  if (sheets.length >= 1) {  
    for(var i = 0;i < sheets.length; i++)
    {
      sheet_names.push(sheets[i].getName());
    }
  }
  return sheet_names;
}

