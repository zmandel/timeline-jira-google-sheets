var g_dateToday = g_blank; //cache the datefor consistency within the same execution
function getDateFormatted(date) {
  if (!date) {
    if (!g_dateToday)
      g_dateToday = Utilities.formatDate(new Date(), Session.getScriptTimeZone(), "yyyy-MM-dd");
    return g_dateToday;
  }
  return Utilities.formatDate(date, Session.getScriptTimeZone(), "yyyy-MM-dd");
}

function getToken() {
  return PropertiesService.getScriptProperties().getProperty("token");
}

function removeTempSheets() {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var sheets = ss.getSheets();
  sheets.forEach(sheet => {
    if (sheet.getName().indexOf("(") > 0)
      ss.deleteSheet(sheet);
  });
}

function createSheet() {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  const nameBase = s_prefixNameSheet + getDateFormatted();
  var name = nameBase;
  var iTry = 2;
  var sheetBefore = null;
  do {
    sheetBefore = ss.getSheetByName(name);
    if (sheetBefore) {
      name = nameBase + ` (${iTry})`;
      iTry++;
    }
  } while (sheetBefore);

  var sheet = ss.insertSheet(s_numSheetsPre);
  sheet.setName(name);
  sheet.getRange(1, 1).setValue(" "); //if done later, a sheets bug shows where it shows help text about using "@"
  SpreadsheetApp.flush();
  return sheet;
}

var t_response = null; //apps script quirks require this to be global (locals get lost in closure stringification)
function setJiraToken() {
  var ui = SpreadsheetApp.getUi();
  t_response = ui.prompt('Admin', 'token:', ui.ButtonSet.OK_CANCEL);

  if (t_response.getSelectedButton() === ui.Button.OK) {
    submitToken(t_response.getResponseText());
    configureMenu();
  }
}

function submitToken(token) {
  if (token)
    PropertiesService.getScriptProperties().setProperty("token", token);
}

function getWeekStartDate(date) {
  const dayOfWeek = date.getDay(); // Get current day of the week (0 for Sunday, 1 for Monday, etc.)
  const difference = (dayOfWeek + (g_daysInWeek - 1)) % g_daysInWeek; // Calculate difference to Monday; adjust for Sunday being 0
  const weekStartDate = new Date(date.setDate(date.getDate() - difference));
  weekStartDate.setHours(0, 0, 0, 0);

  return weekStartDate;
}

function getWeekEndDate(date) {
  const dayOfWeek = date.getDay(); // Get current day of the week (0 for Sunday, 1 for Monday, etc.)
  const difference = (dayOfWeek + (g_daysInWeek - 1)) % g_daysInWeek; // Calculate difference to Monday; adjust for Sunday being 0
  const weekEndDate = new Date(date.setDate(date.getDate() - difference + 6));
  weekEndDate.setHours(0, 0, 0, 0);

  return weekEndDate;
}


function assert(expression) {
  if (!expression) {
    console.error(getStackTrace("assert failed!"));
    throw new Error("bye");
  }
}

const getStackTrace = function (message) {
  let s = `Error: ${message}\n`;
  (new Error()).stack
    .split('\n')
    .forEach((token) => { s += `\t${token.trim()}\n` }
    );
  return s;
}

/**
 * "Class" Toaster
 *
 * From http://stackoverflow.com/a/33552904/1677912
 *
 * Wrapper for Spreadsheet.toast() with support for multi-line messages.
 *
 * Constructor:    new Toaster( message, title, timeoutSeconds );
 *
 * @param message         {String}    Toast message, possibly with newlines ('\n')
 * @param title           {String}    (optional) Toast title
 * @param timeoutSeconds  {Number}    (optional) Duration of display, default 3s
 *
 * @returns {Toaster}                 Toaster instance.
 */
var Toaster = function (message, title, timeoutSeconds) {
  if (typeof message == 'undefined')
    throw new TypeError("missing message");

  this.message = this.parseMessage(message);
  this.title = title || g_blank;
  this.timeoutSeconds = timeoutSeconds || 3;
  this.ss = SpreadsheetApp.getActiveSpreadsheet();
};

/**
 * Display Toaster message using previously set parameters.
 */
Toaster.prototype.display = function () {
  this.ss.toast(this.message, this.title, this.timeoutSeconds);
}

/**
 * This is where the magic happens. Prepares multi-line messages for display.
 *
 * @param {String} msg    Toast message, possibly with newlines ('\n')
 *
 * @returns{String}       Message, ready to display.
 */
Toaster.prototype.parseMessage = function (msg) {
  var maxWidth = 52;             // Approx. number of non-breaking spaces required to span toast popup.
  var knob = 1.85;               // Magical approx. ratio of avg char width : non-breaking space width
  var parsedMessage = g_blank;
  const nbsp = String.fromCharCode(160);
  msg = msg.replace(/ /g, nbsp);
  var lines = msg.split('\n');   // Break lines at newline chars

  // Rebuild message with padded lines
  for (var i = 0; i < lines.length; i++) {
    var len = lines[i].length;
    // Build padding string of non-breaking spaces sandwiched with normal spaces.
    var padding = ' '
      + len < (maxWidth / knob) ?
      Array(Math.floor(maxWidth - (lines[i].length * knob))).join(nbsp) + ' ' : g_blank;
    parsedMessage += lines[i] + padding;
  }
  return parsedMessage;
}