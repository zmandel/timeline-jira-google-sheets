
function listIssues() {
  loadSettings();
  const timeZone = Session.getScriptTimeZone(); // Gets the script's time zone
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var sheet = createSheet();
  sheet.getRange(s_rowsFrozen, 1, 1, s_header.length).setFontSize(g_sizeFontTable).setValues([s_header]);
  var iCol = 1;
  sheet.setColumnWidth(iCol++, ColumnWidths.project);
  sheet.setColumnWidth(iCol++, ColumnWidths.summary);
  sheet.setColumnWidth(iCol++, ColumnWidths.type);
  sheet.setColumnWidth(iCol++, ColumnWidths.key);
  sheet.setColumnWidth(iCol++, ColumnWidths.status);
  sheet.setColumnWidth(iCol++, ColumnWidths.assigned);
  sheet.setColumnWidth(iCol++, ColumnWidths.start);
  sheet.setColumnWidth(iCol++, ColumnWidths.end);
  sheet.setFrozenRows(s_rowsFrozen);
  sheet.setFrozenColumns(s_colNames);

  const rowStart = s_rowsFrozen + 1;
  var rows = [];
  var rows_base = [];
  var colorsAlerts = []; //for rows_base
  var rowsName = [];
  var rowsNameBase = [];
  var colors = [];

  const issues = listJiraIssues();
  const dateNow = getDateFormatted();
  var dateRanges = {
    dateMin: g_blank + dateNow,  //clone
    dateMax: g_blank
  };
  issues.forEach(issue => {
    processIssue(issue, rows, rows_base, rowsName, rowsNameBase, colors, colorsAlerts, dateRanges, dateNow, addRow);
  });

  if (rows.length > 0) {
    sheet.getRange(rowStart, 1, rows.length, s_header.length).setFontSize(g_sizeFontTable).setValues(rows);
    sheet.getRange(rowStart, s_colNames, rows.length, 1).setRichTextValues(rowsName);
    sheet.getRange(rowStart, s_colNames > 1 ? s_colNames - 1 : s_colNames, rows.length, s_colNames).setBackgrounds(colors);
    sheet.getRange(rowStart, s_header.length, rows.length, 1).setBorder(
      null, null, null, true, null, null, "black", SpreadsheetApp.BorderStyle.SOLID);
    if (sheet.getMaxRows() > rows.length + rowStart - 1)
      sheet.deleteRows(rows.length + rowStart, sheet.getMaxRows() - rows.length - (rowStart - 1));
    var rgDates = [];
    if (dateRanges.dateMax) {
      const dateStart = getWeekStartDate(new Date(dateRanges.dateMin));
      const dateLast = getWeekStartDate(new Date(dateRanges.dateMax));
      const dateWeekNow = getWeekStartDate(new Date());
      if (dateLast >= dateStart) {
        const weeks = Math.ceil((dateLast - dateStart + 1) / (1000 * 60 * 60 * 24 * g_daysInWeek));
        if (weeks > 0) {
          const colsMax = sheet.getMaxColumns();
          if (s_header.length + weeks > colsMax)
            sheet.insertColumnsBefore(s_header.length + 1, (s_header.length + weeks) - colsMax);
          else if (s_header.length + weeks < colsMax)
            sheet.deleteColumns(s_header.length + weeks + 1, colsMax - (s_header.length + weeks));
          sheet.setColumnWidths(s_header.length + 1, weeks, g_colwidthPeriod);
          var rangeCols = sheet.getRange(1, s_header.length + 1, 1, weeks);
          var rg = [];
          var rgColors = [];
          for (var dateCur = dateStart, icol = 1; icol <= weeks; icol++, dateCur.setDate(dateCur.getDate() + g_daysInWeek)) {
            rg.push(Utilities.formatDate(dateCur, timeZone, "yyyy") + "\n" + Utilities.formatDate(dateCur, timeZone, "MM-dd"));
            rgDates.push(new Date(dateCur));
            if (dateWeekNow.getTime() == dateCur.getTime())
              rgColors.push(s_colorCurrentWeek);
            else
              rgColors.push(s_colorWhite);
          }
          rangeCols.setFontSize(g_sizeFontWeeks).setHorizontalAlignment("center");
          rangeCols.setValues([rg]);
          rangeCols.setBackgrounds([rgColors]);
        }
      }
    }

    if (rgDates.length > 0) {
      var rgColors = [];
      issues.forEach(issue => {
        if (!issue.bSkip) {
          rgColors.push(generateColorRow(issue, rgDates));
          if (issue.children) {
            issue.children.forEach(issueChild => {
              rgColors.push(generateColorRow(issueChild, rgDates));
            });
          }
        }
      });
      sheet.getRange(rowStart, s_header.length + 1, rows.length, rgDates.length).setBackgrounds(rgColors);
    }
    var rowCur = rowStart;
    issues.forEach(issue => {
      if (!issue.bSkip) {
        rowCur++;
        if (issue.children && issue.children.length > 0) {
          var rangeChildren = sheet.getRange(rowCur, 1, issue.children.length);
          rangeChildren.shiftRowGroupDepth(1);
          rowCur += issue.children.length;
        }
      }
    });
  }

  sheet.getRange(s_rowsFrozen, 1, sheet.getMaxRows() - s_rowsFrozen + 1, sheet.getLastColumn()).createFilter();
  SpreadsheetApp.flush();
  collapseAll(sheet);
  var sheetBaseData = ss.getSheetByName(s_nameBaseGraph);
  sheetBaseData.getRange(s_rangeBase).clear();
  const cRowsBase = sheetBaseData.getMaxRows();
  const rowStartBase = 2;
  const rowEndDelete = cRowsBase - 1;
  if (rowEndDelete >= rowStartBase)
    sheetBaseData.deleteRows(rowStartBase, rowEndDelete + 1 - rowStartBase); //this preserves the graph input range
  const cBase = rows_base.length;
  const cDiffIssues = rows.length - cBase;

  if (cBase > 0 || cDiffIssues > 0) {
    if (sheetBaseData.getMaxRows() != rowStartBase) {
      ss.toast(g_blank, "Warning: graph base broken. Reset input range.");
      Utilities.sleep(3000); //so toast is visible, otherwise the next toast overwrites it
    } else {
      if (cDiffIssues > 0)
        addBaseWarningRow(sheet.getName(), dateRanges, rows_base, rowsNameBase, colorsAlerts, cDiffIssues);
      if (cBase > 0)
        addTopSpanRow(dateRanges, rows_base, rowsNameBase, colorsAlerts);

      //insert "in between" the range of rows 1-2, so that the graph input range auto-adjusts
      sheetBaseData.insertRowsAfter(1, rows_base.length - 1);
    }

    sheetBaseData.getRange(rowStartBase, 1, rows_base.length, rows_base[0].length).setValues(rows_base);
    sheetBaseData.getRange(rowStartBase, s_colNames, rows_base.length, 1).setBackgrounds(colorsAlerts);
    sheetBaseData.getRange(rowStartBase, s_colNames, rows_base.length, 1).setRichTextValues(rowsNameBase);
  }
  ss.toast(g_blank, "Done", 2);
}

function processIssue(issue, rows, rows_base, rowsName, rowsNameBase, colors, colorsAlerts, dateRanges, dateNow, addRow) {
  if (issue.bSkip)
    return;

  if (s_alertWhenNoParent && !issue.fields.parent && issue.fields.issuetype.hierarchyLevel < s_levelParent)
    issue.bAlert = true;
  var dateStart = issue.tm_dateStart;
  var dateEnd = issue.tm_dateEnd;

  if (dateStart && dateStart < dateRanges.dateMin)
    dateRanges.dateMin = dateStart;
  if (dateEnd && dateEnd > dateRanges.dateMax)
    dateRanges.dateMax = dateEnd;

  function hasAlert(issue) {
    const statusName = statusCategory(issue);
    if (statusName != s_statusDone) {
      if (statusName != s_statusDoing) {
        if (issue.tm_dateStart && issue.tm_dateStart < dateNow) {
          return true;
        }
      } else if (true) {
        if (issue.tm_dateEnd && issue.tm_dateEnd < dateNow) {
          return true;
        }
      }
    }
    return false;
  }

  if (issue.children) {
    var bAlert = false;
    issue.children.forEach(issueChild => {
      var dateStartChild = issueChild.tm_dateStart;
      var dateEndChild = issueChild.tm_dateEnd;
      if (hasAlert(issueChild)) {
        issueChild.bAlert = true;
        bAlert = true;
      }

      if (dateStartChild) {
        if (dateStartChild < dateRanges.dateMin)
          dateRanges.dateMin = dateStartChild;
      }

      if (dateEndChild) {
        if (dateEndChild > dateRanges.dateMax)
          dateRanges.dateMax = dateEndChild;
      }
    });

    if (bAlert)
      issue.bAlert = true;
  }

  if (hasAlert(issue)) {
    issue.bAlert = true;
  }

  addRow(rows, rows_base, rowsName, rowsNameBase, colors, colorsAlerts, issue);
  if (issue.children) {
    issue.children.forEach(issueChild => {
      addRow(rows, rows_base, rowsName, rowsNameBase, colors, colorsAlerts, issueChild);
    });
  }
}

function generateColorRow(issue, rgDates) {
  const cColsWeeks = rgDates.length;
  const dateStart = issue.tm_dateStart;
  const dateEnd = issue.tm_dateEnd;
  var rowColors = new Array(cColsWeeks);
  if (dateStart && dateEnd) {
    const dateLeft = getWeekStartDate(new Date(dateStart));
    const dateRight = getWeekEndDate(new Date(dateEnd));
    const hierarchyLevel = issue.fields.issuetype.hierarchyLevel;
    const bTopLevel = hierarchyLevel >= s_levelParent;
    for (var i = 0; i < cColsWeeks; i++) {
      const dateCompare = rgDates[i];
      if (dateCompare >= dateLeft && dateCompare <= dateRight) {
        if (bTopLevel)
          rowColors[i] = s_colorBarParent;
        else
          rowColors[i] = s_colorBarChild;
      }
    }
  }
  return rowColors;
}

function addBaseWarningRow(nameSheet, dateRanges, rows_base, rowsNameBase, colorsAlerts, cDiffIssues) {
  const rowBase = [
    `see sheet '${nameSheet}'`,       //project
    g_blank,                          //summary
    "alert",                          //type
    g_blank,                          //key
    g_blank,                          //status
    g_blank,                          //assigned to
    g_blank,                          //parent
    dateRanges.dateMin,               //start
    dateRanges.dateMax,               //end
  ];
  rows_base.unshift(rowBase);
  colorsAlerts.unshift([s_colorAlertIssue]);
  var textBase = `${g_alertChar}${cDiffIssues} issues without dates`;
  var richValueBase = SpreadsheetApp.newRichTextValue().setText(textBase);
  rowsNameBase.unshift([richValueBase.build()]);
}

function addTopSpanRow(dateRanges, rows_base, rowsNameBase, colorsAlerts) {
  const rowBase = [
    g_blank,               //project
    g_blank,               //summary
    "info",                //type
    g_blank,               //key
    g_blank,               //status
    g_blank,               //assigned to
    g_blank,               //parent
    dateRanges.dateMin,    //start
    dateRanges.dateMax,    //end
  ];
  rows_base.unshift(rowBase);
  colorsAlerts.unshift([s_colorWhite]);
  var richValueBase = SpreadsheetApp.newRichTextValue().setText("Overall Timeline");
  rowsNameBase.unshift([richValueBase.build()]);
}

function addRow(rows, rows_base, rowsName, rowsNameBase, colors, colorsAlerts, issue) {
  var nameParent = g_blank;
  const link = `https://${s_jiraDomain}/browse/${issue.key}`;
  var richValue = SpreadsheetApp.newRichTextValue();
  const prefixShared = (issue.bAlert ? g_alertChar : g_blank);
  const summary = issue.fields.summary;
  var indent = g_blank;
  var prefix = prefixShared;
  const level = issue.fields.issuetype.hierarchyLevel;
  const isTopLevel = (level >= s_levelParent);

  if (isTopLevel)
    nameParent = issue.fields.summary;  //set to self so the native sheets timeline looks good
  else if (issue.fields.parent) {
    indent = "   ";
    prefix = indent + prefix;
    nameParent = issue.fields.parent.fields.summary;
  } else {
    //nothing
  }
  const text = prefix + summary;
  richValue = richValue.setText(text).setLinkUrl(prefix.length, text.length, link);
  rowsName.push([richValue.build()]);

  const row = [
    indent + issue.fields.project.name,           //project
    g_blank,                                      //summary
    issue.fields.issuetype.name,                  //type
    issue.key,                                    //key
    issue.fields.status.name,                     //status
    issue.fields.assignee?.displayName ?? g_blank,//assigned to
    issue.fields[s_propStartDate],                //start
    issue.fields.duedate,                         //end
  ];
  var ret = rows.length;
  rows.push(row);

  if (issue.tm_dateStart && issue.tm_dateEnd) {
    const rowBase = [
      issue.fields.project.name,                    //project
      g_blank,                                      //summary
      issue.fields.issuetype.name,                  //type
      issue.key,                                    //key
      issue.fields.status.name,                     //status
      issue.fields.assignee?.displayName ?? g_blank,//assigned to
      nameParent,                                   //parent (same as parent to make native timelines better)
      issue.tm_dateStart,                           //start
      issue.tm_dateEnd,                             //end
    ];
    rows_base.push(rowBase);
    const textBase = prefixShared + summary;
    var richValueBase = SpreadsheetApp.newRichTextValue().setText(textBase).setLinkUrl(prefixShared.length, textBase.length, link); //without spaces prefixes
    rowsNameBase.push([richValueBase.build()]);
    if (isTopLevel)
      colorsAlerts.push([s_colorParent]);
    else if (issue.bAlert)
      colorsAlerts.push([s_colorAlertIssue]);
    else
      colorsAlerts.push([s_colorWhite]);
  }
  if (isTopLevel)
    colors.push([s_colorParent, s_colorParent]);
  else
    colors.push([s_colorWhite, s_colorWhite]);
  return ret;
}
