/**
 * @return void
 */
function addOpenTrigger() {
  ScriptApp.newTrigger('startUp')
    .forSpreadsheet(getSpreadsheet())
    .onOpen()
    .create();
}

/**
 * @return Spreadsheet
 */
function getSpreadsheet() {
  return SpreadsheetApp.openById(PropertiesService.getScriptProperties().getProperty('spreadsheet_id'));
}

/**
 * @param  String name
 * @return Sheet
 */
function getSheet(name) {
  return getSpreadsheet().getSheetByName(name);
}

/**
 * @return Array
 */
var getOptions = (function() {
  var options = [];

  return function getOptions() {
    var values = getSheet('options').getDataRange().getValues();

    options = [].concat.apply([], values);

    return options;
  }
})();

var SidebarHtml = (function() {
  var html;

  return function SidebarHtml() {
    var tmpl = HtmlService.createTemplateFromFile('autocomplete-options.html');
    html = tmpl.evaluate();
    html.setSandboxMode(HtmlService.SandboxMode.IFRAME)
    .setTitle('登録会社から選択')
    .setWidth(200);

    return html;
  }
})();

function startUp() {
  addMenu();
  generateSidebar();
}

function addMenu() {
  SpreadsheetApp.getUi().createMenu('管理')
    .addItem('社名補完', 'openSidebar')
    .addToUi();
}

function generateSidebar() {
  SidebarHtml();
}

function openSidebar() {
  SpreadsheetApp.getUi().showSidebar(SidebarHtml());
}

/**
 * @param String value
 */
function sendData(value) {
  if ( SpreadsheetApp.getActiveSheet().getSheetName() == 'options' ) {
    Browser.msgBox("You cannot break options source sheet");
  } else {
    SpreadsheetApp.getActiveRange().setValue(value);
  }
}
