//初始化各表格
var ss = SpreadsheetApp.getActiveSpreadsheet();
var Facebook = ss.getSheetByName("Facebook");
var ui = SpreadsheetApp.getUi();

//添加功能菜單
function onOpen() {
  SpreadsheetApp.getUi() // Or DocumentApp or FormApp.
  .createMenu('FTOFS') 
  .addItem('異常賬號提交','showabnAccDialog')
  .addToUi();
}

//添加add-on菜單
//function onOpen(e) {
//  var menu = ui.createAddonMenu(); // Or DocumentApp or FormApp.
//  if (e && e.authMode == ScriptApp.AuthMode.NONE)
//  {
//    // Add a normal menu item (works in all authorization modes).
//    menu.addItem('異常賬號提交','showabnAccDialog');
//  } 
//  else
//  {
//     //Add a menu item based on properties (doesn't work in AuthMode.NONE).
//    var properties = PropertiesService.getDocumentProperties();
//    var workflowStarted = properties.getProperty('workflowStarted');
//    if (workflowStarted)
//    {
//      menu.addItem('異常賬號提交','showabnAccDialog');
//    } 
//    else 
//    {
//      menu.addItem('異常賬號提交','showabnAccDialog');
//    }
//     //Record analytics.
//     //UrlFetchApp.fetch('http://www.example.com/analytics?event=open');
//  }
//  menu.addToUi();
//}

//異常賬號提交對話框
function showabnAccDialog(){
  //當前文件無Facebook表則報錯
  if (Facebook == null)
  {
    ui.alert('當前文件中不存在Facebook賬號表格！\n請在賬號表中使用本插件！');
  }
  else
  {
  var html = HtmlService.createHtmlOutputFromFile('abnormalAccount')
      .setWidth(400)
      .setHeight(200);
  SpreadsheetApp.getUi().showModalDialog(html, '異常賬號提交');
  }
}


//根據給定的表及參數，查找對應行
function find(value){
  var ARange = Facebook.getRange("A2:A");
  var data = ARange.getValues();
  for(var i=0;i<data.length;i++)
  {
    for(var j=0;j<data[i].length;j++)
    {
      if(data[i][j] == value)
        return i+2;
    }
  }
  return null;
}

//獲取賬號詳細信息
function getInfo(accID){
  var Row = find(accID);
  if(Row == null)
  {
    return null;
  }
  else
  {
  var Info = new Array();
  Info.push(Row);
  Info.push(Facebook.getRange(Row, 2).getValue());
  Info.push(Facebook.getRange(Row, 3).getValue());
  Info.push(Facebook.getRange(Row, 5).getValue());
  return Info;
  }
}

//確認提交對話框
function showConfirmAlert(accID,typevalue,note){
  var AccID = accID;
  if (AccID == '')
    {
    //ui.alert('賬號編號不能為空！\n請重新填寫正確的賬號編號！');
    var noAccID = ui.alert('賬號編號不能為空','按“确定”重新输入编号！',ui.ButtonSet.OK);
    if (noAccID == ui.Button.OK)
      showabnAccDialog();
    }
  else
  {
  var Status = typevalue;
  var Note = note;
  var Info = getInfo(accID);
    if(Info == null)
    {
      var AccIDerror = ui.alert('賬號編號錯誤','不存在該賬號\n按“确定”重新输入编号！',ui.ButtonSet.OK);
      if (AccIDerror == ui.Button.OK)
        showabnAccDialog(); 
    }
    else
    {
      var Row = Info[0];
      var User = Info[1];
      var Name = Info[2];
      var Type = Info[3];
      
      var response = ui.alert('請確認需要提交的賬號信息！','編號：'+AccID+'\n使用者：'+User+'\n賬號姓名：'+Name+'\n狀態：'+Status+'\n類型：'+Type+'\n備註：'+Note+'\n\n按“否”重新輸入編號。',ui.ButtonSet.YES_NO); 
      if (response == ui.Button.YES)
        abnormalAccount(AccID,User,Name,Status,Type,Note,Row);
      else 
        showabnAccDialog(); 
    }
  }
}

//提交異常賬號
function abnormalAccount(accID,user,name,status,type,note,row){
  var abnFB = SpreadsheetApp.openById("1WXMf-2VEGxvJLZgLtIJWkYaLUm3cN0-43qTBKsbAvgY");
  var FacebookHost = abnFB.getSheetByName("Facebook異常賬號");
  
  if(FacebookHost == null)
  {
    ui.alert('後臺異常！\n請聯繫IT部。');
  }
  else
  {
  var longdate = new Date();
  var submitdate = Utilities.formatDate(new Date(), "GMT", "yyyy-MM-dd");
  FacebookHost.appendRow([accID,user,name,status,type,note,submitdate]);
  
  Facebook.getRange(row, 4).setValue(status);
  var tempnote = Facebook.getRange(row, 10).getValue();
  if(tempnote == '')
    Facebook.getRange(row, 10).setValue(note);
  else
    Facebook.getRange(row, 10).setValue(note+'，'+tempnote);
   
  ui.alert('提交成功！\n可在異常賬號列表中查詢處理狀態。');
  } 
}