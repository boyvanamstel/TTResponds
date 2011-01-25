/**
 * Copyright 2010, Boy van Amstel
 * All rights reserved.
 */

var NAME = 'TTResponds';
var CONFIGSHEET = 'TTRespondsConfig';
var _activeApp = SpreadsheetApp.getActive();
var _formSheet = _activeApp.getSheets()[0];

var _config = new Object();
var _params = [
        ['from', 'From (name)', Session.getUser().getUserLoginId()],
        ['subject', 'Subject', 'Thanks for your response'],
        ['body', 'Body', 'Hi! Thanks! Bye!'],
        ['emailColumn', 'Email column', 3]
      ];

//var _emailColumn = 2; // Starts at zero - deprecated
var _quota = 0;

function onFormSubmit(e) {

  send();

}

function getLastEmailAddress() {
  loadConfig()
  var dataRange = _formSheet.getRange(_formSheet.getLastRow(), 1, 1, _config.emailColumn);
  return dataRange.getValues()[0][_config.emailColumn -1];
}

function showMsg(msg) {
  Browser.msgBox(msg);
}

function onOpen() {
  createMenu();
  updateQuota();
  if(checkConfigExists()) {
    _config.sheet = _activeApp.getSheetByName(CONFIGSHEET);
  } else {
    createConfig();
  }
}

function onInstall() {
  //onOpen();
  createMenu();
}

function checkConfigExists() {
  if(_activeApp.getSheetByName(CONFIGSHEET)) return true;
  return false;
}

function createConfig() {
  if(checkConfigExists()) return;
  
  updateQuota();
  
   _config.sheet =  _activeApp.insertSheet(CONFIGSHEET);
  
  _config.sheet.setColumnWidth(1, 200);
  _config.sheet.setColumnWidth(2, 200);
  
  for(var i = 0; i < _params.length; i++) {
    _config[_params[i][1]] = _params[i][2];
  }
   
   
  // Set title 'n stuff
  _config.sheet.getRange('A1:A1').setValue('TTResponds configuration');
  _config.sheet.getRange('A1:A1').setFontSize(18);
  
  values = new Array();
  values.push(['Quota', _quota]);
  values.push(['','']);

  for(var i = 0; i < _params.length; i++) {
    values.push([_params[i][1], _params[i][2]]);
  }
  
  _config.sheet.getRange('A3:B' + (values.length + 2)).setValues(values);
  _config.sheet.getRange('A5:A' + (values.length + 2)).setBackgroundColor("#eee");
}

function createMenu() {
  var subMenus = [{name:'About', functionName:'about'}, {name:'Create config', functionName:'createConfig'}]
  _activeApp.addMenu(NAME, subMenus);
}

function updateQuota() {
  _quota = MailApp.getRemainingDailyQuota();
  _config.quota = _quota;
  if(checkConfigExists()) {
    loadConfig();
    _config.sheet.getRange('B3:B3').setValue(_quota);
    }
}

function loadConfig() {
  _config = { sheet: _activeApp.getSheetByName(CONFIGSHEET) };
  
  if(_config.sheet) {
    var values = _config.sheet.getRange('B5:B' + (_params.length + 5)).getValues();

    for(var i = 0; i < _params.length; i++) {
      //showMsg(_params[i][0] + ': ' + values[i]);
      _config[_params[i][0]] = values[i].toString();
    }
  }
}

function send() {
  loadConfig();

  if(_formSheet.getLastRow() > 1) {
    MailApp.sendEmail(
      getLastEmailAddress(), 
      _config.subject, 
      _config.body,
      { name: _config.from }
      );  

     updateQuota();
   }
}

function about() {
  showMsg('Created by Boy van Amstel / Tam Tam, http://boyvanamstel.nl');
}
​