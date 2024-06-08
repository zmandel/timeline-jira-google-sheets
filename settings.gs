const s_alertWhenNoParent = true; //useful in the Epics & Stories mode so issues without epics get alerted
const s_prefixNameSheet = ""; //adds a prefix to the timeline sheets
const s_rowsFrozen = 2;
const s_numSheetsPre = 3; //New timeline sheets are inserted after these (Settings, base_graph, graph) 
const s_colNames = 2; //can be 1 or 2 depending on how many you want to freeze
const s_rangeBase = "A2:I"; //Warning: this changes if you add more columns. Includes the first week column
const s_nameBaseGraph = "base_graph";
const s_colorWhite = "#FFFFFF";
const g_alertChar = "❗";
const g_sizeFontTable = 8;
const g_sizeFontWeeks = 6;
const g_secWaitShort = 3;
const g_secWaitLong = 7;


//Default column widths
//It gets changed to the columns widths in the last timeline sheet, if any.
//Note: order defines s_header order
var ColumnWidths = {
  project: 110,
  summary: 220,
  type: 50,
  key: 95,
  status: 95,
  assigned: 95,
  start: 65,
  end: 65,
};

var g_colwidthPeriod = 32;

const Strings = {
  sheetSettings: "Settings",
  modeStories: "Stories & Subtasks",
};


const s_header = Object.keys(ColumnWidths); //header for the "day" sheets
const g_daysInWeek = 7;

const Levels = {
  Epic: 1,
  Story: 0,
  Subtask: -1,
};

const g_blank = "";
var s_levelParent = Levels.Epic;
var s_propStartDate = null;
var s_email = null;
var s_jiraDomain = null;
var s_filterId = null;
var s_statusDoing = null;
var s_statusDone = null;
var s_colorAlertIssue = null;
var s_colorCurrentWeek = null;
var s_colorBarParent = null;
var s_colorBarChild = null;
var s_colorParent = null;

function loadSettings() {
  const rowLast = 12;
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var sheet = ss.getSheetByName(Strings.sheetSettings);
  var range = sheet.getRange(`B1:B${rowLast}`);
  const values = range.getValues();
  const colors = range.getBackgrounds();
  var iRow = 0;
  s_mode = values[iRow++][0];

  if (s_mode == Strings.modeStories)
    s_levelParent = Levels.Story;
  else
    s_levelParent = Levels.Epic;

  function nextValue() {
    return values[iRow++][0];
  }

  function nextColor() {
    return colors[iRow++][0];
  }

  s_jiraDomain = nextValue();
  s_propStartDate = nextValue();
  s_filterId = String(nextValue());
  s_statusDoing = nextValue().toLowerCase();
  s_statusDone = nextValue().toLowerCase();
  s_colorCurrentWeek = nextColor();
  s_colorBarParent = nextColor();
  s_colorBarChild = nextColor();
  s_colorParent = nextColor();
  s_colorAlertIssue = nextColor();
  s_email = nextValue();
  assert(iRow == rowLast);

  loadDefaultStyles(ss);
}

function loadDefaultStyles(ss) {
  var sheets = ss.getSheets();
  const pattern = /.*-.*-.*/;
  const cSheets = sheets.length;
  var name = "";
  for (var iSheet = 0; iSheet < cSheets; iSheet++) {
    const nameCur = sheets[iSheet].getName();
    if (pattern.test(nameCur)) {
      name = nameCur;
      break;
    }
  }

  if (name) {
    var sheet = ss.getSheetByName(name);
    var iCol = 1;
    ColumnWidths.project = sheet.getColumnWidth(iCol++);
    ColumnWidths.summary = sheet.getColumnWidth(iCol++);
    ColumnWidths.type = sheet.getColumnWidth(iCol++);
    ColumnWidths.key = sheet.getColumnWidth(iCol++);
    ColumnWidths.status = sheet.getColumnWidth(iCol++);
    ColumnWidths.assigned = sheet.getColumnWidth(iCol++);
    ColumnWidths.start = sheet.getColumnWidth(iCol++);
    ColumnWidths.end = sheet.getColumnWidth(iCol++);
    g_colwidthPeriod = sheet.getColumnWidth(iCol++);
  }
}


function saveSettings() {
  var ss = SpreadsheetApp.getActiveSpreadsheet();

  const saveStages = {
    error: "Error",
    saving: "Saving",
    saved: "Saved",
  };

  const apiToken = PropertiesService.getScriptProperties().getProperty("token");
  if (!apiToken) {
    ss.toast("Configure the Jira API token from the Admin menu.", saveStages.error, g_secWaitLong);
    return;
  }


  var sheet = ss.getSheetByName(Strings.sheetSettings);
  loadSettings();
  if (!s_mode || !s_propStartDate || !s_jiraDomain || !s_filterId || !s_statusDoing || !s_statusDone || !s_email) {
    ss.toast("Fill all the configuration cells.", saveStages.error, g_secWaitLong);
    return;
  }
  ss.toast("Testing Jira connection...", saveStages.saving, 50);
  const issues = listJiraIssues();
  ss.toast(`Filter returned ${issues.length} issues.`, saveStages.saving, g_secWaitShort);
  Utilities.sleep(g_secWaitShort * 1000);
  var bFoundDoing = false;
  var bFoundDone = false;

  for (const key in g_mapStatusCategories) {
    const val = g_mapStatusCategories[key];
    if (val == s_statusDoing)
      bFoundDoing = true;
    else if (val == s_statusDone)
      bFoundDone = true;
  }

  var bShowStatuses = false;
  function toastNotFound(status) {
    ss.toast(`status category '${status}' not found.`, saveStages.error, g_secWaitShort);
    bShowStatuses = true;
  }

  if (!bFoundDoing)
    toastNotFound(s_statusDoing);
  else if (!bFoundDone)
    toastNotFound(s_statusDone);

  if (bShowStatuses) {
    Utilities.sleep(g_secWaitShort * 1000);
    const uniqueValues = new Set();

    for (const key in g_mapStatusCategories)
      uniqueValues.add(g_mapStatusCategories[key]);

    const sortedUniqueValues = Array.from(uniqueValues).sort();
    new Toaster(`status should be one of:\n${sortedUniqueValues.join("\n")}`, saveStages.error, g_secWaitLong).display();
  } else {
    sheet.hideSheet();
    new Toaster("View settings from the menu:\nJira\n ⤷Admin\n    ⤷Settings & Help", saveStages.saved, g_secWaitLong).display();
  }
}

function viewSettings() {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var sheet = ss.getSheetByName(Strings.sheetSettings);
  sheet.showSheet();
  sheet.activate();
  if (!getToken()) {
    new Toaster("Follow the instructions to generate and save your Jira API token.", g_secWaitLong).display();
  }

}

