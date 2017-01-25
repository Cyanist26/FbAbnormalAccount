//��ʼ�������
var ss = SpreadsheetApp.getActiveSpreadsheet();
var Facebook = ss.getSheetByName("Facebook");
var ui = SpreadsheetApp.getUi();

//��ӹ��ܲˆ�
function onOpen() {
  SpreadsheetApp.getUi() // Or DocumentApp or FormApp.
  .createMenu('FTOFS') 
  .addItem('�����~̖�ύ','showabnAccDialog')
  .addToUi();
}

//���add-on�ˆ�
//function onOpen(e) {
//  var menu = ui.createAddonMenu(); // Or DocumentApp or FormApp.
//  if (e && e.authMode == ScriptApp.AuthMode.NONE)
//  {
//    // Add a normal menu item (works in all authorization modes).
//    menu.addItem('�����~̖�ύ','showabnAccDialog');
//  } 
//  else
//  {
//     //Add a menu item based on properties (doesn't work in AuthMode.NONE).
//    var properties = PropertiesService.getDocumentProperties();
//    var workflowStarted = properties.getProperty('workflowStarted');
//    if (workflowStarted)
//    {
//      menu.addItem('�����~̖�ύ','showabnAccDialog');
//    } 
//    else 
//    {
//      menu.addItem('�����~̖�ύ','showabnAccDialog');
//    }
//     //Record analytics.
//     //UrlFetchApp.fetch('http://www.example.com/analytics?event=open');
//  }
//  menu.addToUi();
//}

//�����~̖�ύ��Ԓ��
function showabnAccDialog(){
  //��ǰ�ļ��oFacebook��t���e
  if (Facebook == null)
  {
    ui.alert('��ǰ�ļ��в�����Facebook�~̖���\nՈ���~̖����ʹ�ñ������');
  }
  else
  {
  var html = HtmlService.createHtmlOutputFromFile('abnormalAccount')
      .setWidth(400)
      .setHeight(200);
  SpreadsheetApp.getUi().showModalDialog(html, '�����~̖�ύ');
  }
}


//�����o���ı����������Ҍ�����
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

//�@ȡ�~̖Ԕ����Ϣ
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

//�_�J�ύ��Ԓ��
function showConfirmAlert(accID,typevalue,note){
  var AccID = accID;
  if (AccID == '')
    {
    //ui.alert('�~̖��̖���ܞ�գ�\nՈ��������_���~̖��̖��');
    var noAccID = ui.alert('�~̖��̖���ܞ��','����ȷ�������������ţ�',ui.ButtonSet.OK);
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
      var AccIDerror = ui.alert('�~̖��̖�e�`','������ԓ�~̖\n����ȷ�������������ţ�',ui.ButtonSet.OK);
      if (AccIDerror == ui.Button.OK)
        showabnAccDialog(); 
    }
    else
    {
      var Row = Info[0];
      var User = Info[1];
      var Name = Info[2];
      var Type = Info[3];
      
      var response = ui.alert('Ո�_�J��Ҫ�ύ���~̖��Ϣ��','��̖��'+AccID+'\nʹ���ߣ�'+User+'\n�~̖������'+Name+'\n��B��'+Status+'\n��ͣ�'+Type+'\n���]��'+Note+'\n\n����������ݔ�뾎̖��',ui.ButtonSet.YES_NO); 
      if (response == ui.Button.YES)
        abnormalAccount(AccID,User,Name,Status,Type,Note,Row);
      else 
        showabnAccDialog(); 
    }
  }
}

//�ύ�����~̖
function abnormalAccount(accID,user,name,status,type,note,row){
  var abnFB = SpreadsheetApp.openById("1WXMf-2VEGxvJLZgLtIJWkYaLUm3cN0-43qTBKsbAvgY");
  var FacebookHost = abnFB.getSheetByName("Facebook�����~̖");
  
  if(FacebookHost == null)
  {
    ui.alert('���_������\nՈ�MIT����');
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
    Facebook.getRange(row, 10).setValue(note+'��'+tempnote);
   
  ui.alert('�ύ�ɹ���\n���ڮ����~̖�б��в�ԃ̎���B��');
  } 
}