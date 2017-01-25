var ss = SpreadsheetApp.getActiveSpreadsheet();
var mainSheet = ss.getSheetByName("手機號碼情況總覽");
var rechargeRecord = ss.getSheetByName("充值記錄");
var ui = SpreadsheetApp.getUi();

function onOpen() {
  SpreadsheetApp.getUi() // Or DocumentApp or FormApp.
  .createMenu('號碼管理') 
  .addItem('新增充值記錄','showRecordDialog')
  .addToUi();
}

function showRecordDialog(){
  //初始化失敗報錯
  if (mainSheet == null || rechargeRecord == null)
  {
    ui.alert('初始化失敗！\n請檢查表格名設置！');
  }
  else
  {
  var html = HtmlService.createHtmlOutputFromFile('rechargeRecord')
      .setWidth(400)
      .setHeight(300);
  SpreadsheetApp.getUi().showModalDialog(html, '新增充值記錄');
  }
}

//根據給定的表及參數，查找對應行
function find(value){
  var ARange = mainSheet.getRange("A4:A");
  var data = ARange.getValues();
  for(var i=0;i<data.length;i++)
  {
    for(var j=0;j<data[i].length;j++)
    {
      if(data[i][j] == value)
        return i+4;
    }
  }
  return null;
}

//獲取号码詳細信息
function getInfo(numID){
  var Row = find(numID);
  if(Row == null)
  {
    return null;
  }
  else
  {
  var Info = new Array();
  Info[0] = (Row);
  Info[1] = mainSheet.getRange(Row, 2).getValue();
  Info[2] = mainSheet.getRange(Row, 3).getValue();
  return Info;
  }
}

function addRecord(numID,code,recDate,expDate,amount,balance){
  var Info = getInfo(numID);
  if(Info == null)
    {
      var AccIDerror = ui.alert('號碼編號錯誤','不存在該號碼\n按“确定”重新输入编号！',ui.ButtonSet.OK);
      if (AccIDerror == ui.Button.OK)
        showRecordDialog(); 
    }
    else
    {
      var Row = Info[0];
      var Num = Info[1];
      var Operator = Info[2];
      
      rechargeRecord.appendRow([numID,Num,Operator,code,recDate,expDate,amount,balance]);
      mainSheet.getRange(Row, 8).setValue('=IF((DATEVALUE("'+expDate+'")-TODAY())>=5,"'+expDate+'","5天內到期！")');
    }
}