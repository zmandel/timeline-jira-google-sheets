function onOpen() {
  configureMenu();
}

function configureMenu() {
  const apiToken = getToken();
  var ui = SpreadsheetApp.getUi();
  var obj = ui.createMenu('Jira')
    .addItem('Create timeline', 'listIssues')
    .addItem('Collape All', 'collapseAll')
    .addItem('Expand All', 'expandAll');
  var menuAdmin = ui.createMenu('Admin');
  menuAdmin.addItem('Settings & Help', 'viewSettings');
  if (!apiToken)
    menuAdmin.addItem('set JIRA Token', 'setJiraToken');
  obj.addSubMenu(menuAdmin).addToUi();
}

function collapseAll(sheet) {
  try {
    if (!sheet)
      sheet = SpreadsheetApp.getActiveSpreadsheet().getActiveSheet();
    if (sheet.getLastRow() > 0)
      sheet.collapseAllRowGroups();
  } catch (e) {
    //nothing
  }
}

function expandAll(sheet) {
  try {
    if (!sheet)
      sheet = SpreadsheetApp.getActiveSpreadsheet().getActiveSheet();
    if (sheet.getLastRow() > 0)
      sheet.expandAllRowGroups();
  } catch (e) {
    //nothing
  }
}

