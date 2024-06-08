var g_mapStatusCategories = null;

function listJiraIssues() {
  if (!s_email)
    loadSettings();
  g_mapStatusCategories = getJiraStatuses();
  const issues = getIssues();
  issues.sort(function (i1, i2) {
    var result = i2.fields.issuetype.hierarchyLevel - i1.fields.issuetype.hierarchyLevel;
    //just place parents at the beginning
    return result;
  });

  var parents = {};
  issues.forEach(issue => {
    const level = issue.fields.issuetype.hierarchyLevel;
    if (level > s_levelParent) {
      issue.bSkip = true; //filters out issues over the two levels
    }
    else if (level == s_levelParent) {
      parents[issue.key] = issue;
      issue.children = [];
    } else if (issue.fields.parent) {
      var issueParent = parents[issue.fields.parent.key];
      issue.bSkip = true; //this also filters out issues under the two levels when the parent is not found
      if (issueParent)
        issueParent.children.push(issue);
    }
  });

  //standarize start/end dates in issue.tm_dateStart, issue.tm_dateEnd
  issues.forEach(issue => {
    processJiraIssue(issue);
  });

  //sort by start date
  sortIssues(issues);
  return issues;
}

function sortIssues(issues) {
  const dateEOT = getDateFormatted(new Date(2100, 2, 1));
  function sort(i1, i2) {
    var dateStart1 = i1.tm_dateStart;
    var dateStart2 = i2.tm_dateStart;

    if (!dateStart1)
      dateStart1 = dateEOT;

    if (!dateStart2)
      dateStart2 = dateEOT;
    return dateStart1.localeCompare(dateStart2);
  }

  issues.sort(sort);

  issues.forEach(issue => {
    if (!issue.bSkip && issue.children) {
      issue.children.sort(sort);
    }
  });
}



function processJiraIssue(issue) {
  if (issue.bSkip)
    return;
  var dateStart = issue.fields[s_propStartDate];
  var dateEnd = issue.fields.duedate;
  const dateToday = getDateFormatted();

  if (!dateStart)
    dateStart = dateEnd;

  if (!dateEnd)
    dateEnd = dateStart;

  issue.tm_dateStart = dateStart;
  issue.tm_dateEnd = dateEnd;

  if (issue.children) {
    if (issue.children.length == 0) {
      //no children
      const level = issue.fields.issuetype.hierarchyLevel;
      if (level >= Levels.Epic)
        issue.bSkip = true;
    } else {
      issue.children.forEach(issueChild => {
        var dateStartChild = issueChild.fields[s_propStartDate];
        var dateEndChild = issueChild.fields.duedate;
        if (!dateStartChild)
          dateStartChild = dateEndChild;

        if (!dateEndChild)
          dateEndChild = dateStartChild;

        if (dateStartChild && isStatusDoing(issueChild) && dateStartChild > dateToday)
          dateStartChild = dateToday; //not perfect but less misleading

        if (dateEndChild && isStatusDone(issueChild) && dateEndChild > dateToday)
          dateEndChild = dateToday;   //not perfect but less misleading

        if (dateStartChild) {
          if (!dateStart || dateStartChild < dateStart) {
            dateStart = dateStartChild;
            issue.tm_dateStart = dateStart;
          }
        }

        if (dateEndChild) {
          if (!dateEnd || dateEndChild > dateEnd) {
            dateEnd = dateEndChild;
            issue.tm_dateEnd = dateEnd;
          }
        }
        issueChild.tm_dateStart = dateStartChild;
        issueChild.tm_dateEnd = dateEndChild;
      });
    }
  }
}

function getIssues() {
  const apiToken = PropertiesService.getScriptProperties().getProperty("token");
  const fields = `project,key,summary,status,duedate,issuetype,assignee,parent,${s_propStartDate}`;
  const maxResults = 100; //jira doenst allow more
  const apiUrlBase = `https://${s_jiraDomain}/rest/api/2/search?jql=filter=${s_filterId}&fields=${fields}&maxResults=${maxResults}&startAt=`;

  const headers = {
    "Authorization": `Basic ${Utilities.base64Encode(`${s_email}:${apiToken}`)}`,
    "Accept": "application/json"
  };

  const options = {
    "method": "get",
    "headers": headers,
    "muteHttpExceptions": true
  };

  var bStop = false;
  var startAt = 0;
  var issues = [];


  do {
    var response = UrlFetchApp.fetch(apiUrlBase + startAt, options);
    const json = response.getContentText();
    const data = JSON.parse(json);

    if (!data)
      break;

    if (data.errorMessages) {
      new Toaster(data.errorMessages[0],"Error", g_secWaitLong).display();
      console.error(data.errorMessages[0]);
      break;
    }
    if (data.issues) {
      for (var i = 0; i < data.issues.length; i++) {
        const issue = data.issues[i];
        issues.push(issue);

      }
      if (data.issues.length > 0) {
        startAt += data.issues.length;
      } else {
        bStop = true;
      }
    }
  } while (!bStop);

  return issues;
}


function getJiraStatuses() {
  assert(s_email);
  const apiToken = PropertiesService.getScriptProperties().getProperty("token");
  const apiUrlBase = `https://${s_jiraDomain}/rest/api/3/status`;

  var options = {
    'method': 'get',
    'contentType': 'application/json',
    'headers': {
      'Authorization': `Basic ${Utilities.base64Encode(`${s_email}:${apiToken}`)}`,
    },
    'muteHttpExceptions': true
  };

  var response = UrlFetchApp.fetch(apiUrlBase, options);
  var json = JSON.parse(response.getContentText());
  var map = {};
  if (response.getResponseCode() == 200) {
    json.forEach(function (status) {
      map[status.name.toLowerCase()] = status.untranslatedName.toLowerCase();
    });
    return map;
  } else {
    throw new Error(response.getContentText());
  }
}

function statusCategory(issue) {
  const status = issue.fields.status.statusCategory.name.toLowerCase();
  return g_mapStatusCategories[status] ?? status;
}

function isStatusDone(issue) {
  return (statusCategory(issue) == s_statusDone);
}

function isStatusDoing(issue) {
  return (statusCategory(issue) == s_statusDoing);
}
