function getSetDays(){
  var staffCalId = 'vsa.edu.hk_naek8tnu54moqgqfbef6vr85bc@group.calendar.google.com';
  var start = new Date('1-Aug-2025');
  var end = new Date('30-Jun-2026');
  var events = CalendarApp.getCalendarById(staffCalId).getEvents(start, end);
  var daylist = ["Day 1", "Day 2", "Day 3", "Day 4", "Day 5", "Day 6", "Day 7", "Day 8", "Day 9", "Day 10", "Day 0"];
  var outRows = [];
  var tz = Session.getScriptTimeZone() || 'Asia/Hong_Kong';

  for (var i = 0; i < events.length; i++){
    var ev = events[i];
    if (!ev.isAllDayEvent()) continue;
    var title = (ev.getTitle() || '').toString().trim();

    // Find the first day token from daylist that appears anywhere in the title
    var matched = null;
    for (var j = 0; j < daylist.length; j++){
      if (title.indexOf(daylist[j]) !== -1) { matched = daylist[j]; break; }
    }
    if (!matched) continue;

    var evStart = ev.getStartTime();
    var y = parseInt(Utilities.formatDate(evStart, tz, 'yyyy'), 10);
    var m = parseInt(Utilities.formatDate(evStart, tz, 'MM'), 10);
    var d = parseInt(Utilities.formatDate(evStart, tz, 'dd'), 10);
    var localMidnight = new Date(y, m - 1, d, 0, 0, 0);

    // We'll populate: A=matched, B=localMidnight, C=formula (set later), D=full title, E=formula (set later)
    outRows.push([matched, localMidnight, '', title, '']);
  }

  var sheet = SpreadsheetApp.getActiveSheet();
  // Clear previous rows from row 2 downward in the first five columns to avoid leftover data
  try {
    var maxRows = sheet.getMaxRows();
    if (maxRows > 1) sheet.getRange(2, 1, maxRows - 1, 5).clearContent();
  } catch (e) { /* non-fatal */ }

  if (outRows.length > 0) {
    sheet.getRange(2, 1, outRows.length, 5).setValues(outRows);

    // Set formulas for columns C and E with dynamic row numbers
    for (var r = 0; r < outRows.length; r++){
      var rowNum = 2 + r;
      var formulaC = '=IF(WEEKDAY(B' + rowNum + ')=4,3,IF(E' + rowNum + ',2,1))';
      var formulaE = '=IFERROR(SEARCH("core",LOWER(D' + rowNum + '))>0,FALSE)';
      sheet.getRange(rowNum, 3).setFormula(formulaC);
      sheet.getRange(rowNum, 5).setFormula(formulaE);
    }

    // Apply a date/time number format to column B rows 2+ so the local-midnight shows clearly
    try { sheet.getRange(2, 2, outRows.length, 1).setNumberFormat('dd-MMM (ddd) HH:mm'); } catch (e) { /* non-fatal */ }
  }
}

function include(filename) {
  return HtmlService.createHtmlOutputFromFile(filename).getContent();
}

function doGet(e) {
  return HtmlService.createTemplateFromFile('Index').evaluate()
    .setTitle('Schoogle')
    .setXFrameOptionsMode(HtmlService.XFrameOptionsMode.ALLOWALL);
}

function getInitialData() {
  var email = Session.getActiveUser().getEmail() || Session.getEffectiveUser().getEmail();
  var ss = SpreadsheetApp.getActive();
  var periods = [];
  var userTimetable = [];
  var defaultFrom = '';
  var defaultTo = '';

  // PeriodNames: Columns: periodID, Display Text
  var periodSheet = ss.getSheetByName('PeriodNames');
  if (periodSheet) {
    var pVals = periodSheet.getDataRange().getValues();
    for (var i = 1; i < pVals.length; i++) {
      var pid = (pVals[i][0] || '').toString().trim();
      var ptxt = (pVals[i][1] || '').toString().trim();
      if (pid && ptxt) periods.push({ id: pid, text: ptxt });
    }
  }

  // TIMETABLE_MASTER: EMAIL(A), teacher Name(B), period Name(C), class/title(D), day(E), time(F), room/location(G)
  var ttSheet = ss.getSheetByName('TIMETABLE_MASTER');
  if (ttSheet) {
    var tVals = ttSheet.getDataRange().getValues();
    // Collect user timetable if email present
    for (var r = 1; r < tVals.length; r++) {
      var row = tVals[r];
      if (email && (row[0] || '').toString().trim().toLowerCase() === email.toLowerCase()) {
        userTimetable.push({
          day: (row[4] || '').toString().trim(),
          periodId: (row[2] || '').toString().trim(),
          title: (row[3] || '').toString().trim(),
          location: (row[6] || '').toString().trim()
        });
      }
    }
  }

  // Build a teacher list (email + display name) from TIMETABLE_MASTER for client-side selection
  var teachers = [];
  try {
    if (ttSheet) {
      var tVals2 = ttSheet.getDataRange().getValues();
      var seen = {};
      for (var tr = 1; tr < tVals2.length; tr++) {
        var trow = tVals2[tr];
        var te = (trow[0] || '').toString().trim();
        var tn = (trow[1] || '').toString().trim();
        if (te && !seen[te]) { seen[te] = true; teachers.push({ email: te, name: tn || te }); }
      }
      // sort alphabetically by display name
      teachers.sort(function(a,b){ return (a.name || '').toString().localeCompare((b.name||'').toString()); });
    }
  } catch (e) { teachers = []; }

  // Determine default date range from Days sheet (col B)
  var daysSheet = ss.getSheetByName('Days');
  if (daysSheet) {
    var dVals = daysSheet.getDataRange().getValues();
    var minD = null, maxD = null;
    for (var i2 = 1; i2 < dVals.length; i2++) {
      var dCell = dVals[i2][1];
      if (!dCell) continue;
      var dObj = normalizeToDate(dCell);
      if (!minD || dObj.getTime() < minD.getTime()) minD = dObj;
      if (!maxD || dObj.getTime() > maxD.getTime()) maxD = dObj;
    }
    if (minD) defaultFrom = Utilities.formatDate(minD, Session.getScriptTimeZone() || 'Asia/Hong_Kong', 'yyyy-MM-dd');
    if (maxD) defaultTo = Utilities.formatDate(maxD, Session.getScriptTimeZone() || 'Asia/Hong_Kong', 'yyyy-MM-dd');
  }

  return { periods: periods, userTimetable: userTimetable, userEmail: email, defaultFrom: defaultFrom, defaultTo: defaultTo, teachers: teachers };
}

// Return a list of teachers (email + display name)
function getTeachers() {
  var ss = SpreadsheetApp.getActive();
  var ttSheet = ss.getSheetByName('TIMETABLE_MASTER');
  var out = [];
  if (!ttSheet) return out;
  var vals = ttSheet.getDataRange().getValues();
  var seen = {};
  for (var i = 1; i < vals.length; i++) {
    var row = vals[i];
    var email = (row[0] || '').toString().trim();
    var name = (row[1] || '').toString().trim();
    if (!email) continue;
    if (seen[email]) continue;
    seen[email] = true;
    out.push({ email: email, name: name || email });
  }
  out.sort(function(a,b){ return (a.name || '').toString().localeCompare((b.name||'').toString()); });
  return out;
}

// Get all unique classes from TIMETABLE_MASTER sheet (column D)
function getClasses() {
  var ss = SpreadsheetApp.getActive();
  var ttSheet = ss.getSheetByName('TIMETABLE_MASTER');
  var out = [];
  if (!ttSheet) return out;
  var vals = ttSheet.getDataRange().getValues();
  var seen = {};
  for (var i = 1; i < vals.length; i++) {
    var row = vals[i];
    var className = (row[3] || '').toString().trim(); // Column D (index 3)
    if (!className) continue;
    if (seen[className]) continue;
    // Skip classes that start with "Duty" or "MSMEET"
    if (className.toLowerCase().startsWith('duty') || className.toLowerCase().startsWith('msmeet')) continue;
    seen[className] = true;
    out.push(className);
  }
  // Sort alphabetically
  out.sort(function(a, b) { return a.localeCompare(b); });
  return out;
}

// Return timetable rows for a given teacher email
function getTeacherTimetable(teacherEmail) {
  var out = [];
  if (!teacherEmail) return out;
  var ss = SpreadsheetApp.getActive();
  var ttSheet = ss.getSheetByName('TIMETABLE_MASTER');
  if (!ttSheet) return out;
  var vals = ttSheet.getDataRange().getValues();
  for (var i = 1; i < vals.length; i++) {
    var row = vals[i];
    var email = (row[0] || '').toString().trim();
    if (!email || email.toLowerCase() !== teacherEmail.toLowerCase()) continue;
    out.push({ day: (row[4] || '').toString().trim(), periodId: (row[2] || '').toString().trim(), title: (row[3] || '').toString().trim(), location: (row[6] || '').toString().trim() });
  }
  return out;
}

// Return timetable rows for a given class name
function getClassTimetable(className) {
  var out = [];
  if (!className) return out;
  var ss = SpreadsheetApp.getActive();
  var ttSheet = ss.getSheetByName('TIMETABLE_MASTER');
  if (!ttSheet) return out;
  var vals = ttSheet.getDataRange().getValues();
  for (var i = 1; i < vals.length; i++) {
    var row = vals[i];
    var title = (row[3] || '').toString().trim(); // Column D (index 3)
    if (!title || title.toLowerCase() !== className.toLowerCase()) continue;
    out.push({ day: (row[4] || '').toString().trim(), periodId: (row[2] || '').toString().trim(), title: (row[3] || '').toString().trim(), location: (row[6] || '').toString().trim() });
  }
  return out;
}

function getCalendarList() {
  // Use Advanced Calendar Service to get calendars with write access
  var out = [];
  var pageToken;
  do {
    var resp = Calendar.CalendarList.list({ pageToken: pageToken, maxResults: 250 });
    var items = (resp && resp.items) || [];
    items.forEach(function (cal) {
      if (cal && (cal.accessRole === 'owner' || cal.accessRole === 'writer')) {
        out.push({ id: cal.id, name: cal.summary || cal.id });
      }
    });
    pageToken = resp.nextPageToken;
  } while (pageToken);
  // Primary first
  out.sort(function (a, b) {
    if (a.id === 'primary') return -1;
    if (b.id === 'primary') return 1;
    return a.name.localeCompare(b.name);
  });
  return out;
}

function createNewCalendar(calendarName) {
  var tz = Session.getScriptTimeZone() || 'Asia/Hong_Kong';
  var created = Calendar.Calendars.insert({ summary: calendarName, timeZone: tz });
  return created.id;
}

function generatePreviewEvents(formData) {
  var ss = SpreadsheetApp.getActive();


  // Days: Columns: Day, Date, DayType, CycleNum
  var daysSheet = ss.getSheetByName('Days');
  var dayRows = daysSheet ? daysSheet.getDataRange().getValues() : [];
  var days = [];
  for (var i = 1; i < dayRows.length; i++) {
    var d = dayRows[i];
    days.push({
      Day: (d[0] || '').toString().trim(),
      Date: d[1],
      DayType: d[2].toString().trim(), //(d[2] || '').toString().trim(),
      CycleNum: (d[5] || '').toString().trim()
    });
  } 


  // DayTimes: Columns: DayType, PeriodID, START TIME, END TIME
  var dtSheet = ss.getSheetByName('DayTimes');
  var dtRows = dtSheet ? dtSheet.getDataRange().getValues() : [];
  var dayTimesMap = {}; // key: DayType|PeriodID -> {start: Date|String, end: Date|String}
  for (var j = 1; j < dtRows.length; j++) {
    var row = dtRows[j];
    var key = [(row[0] || '').toString().trim(), (row[1] || '').toString().trim()].join('|');
    dayTimesMap[key] = { start: row[2], end: row[3] };
  }


  // Determine date window
  var payload = Array.isArray(formData) ? { items: formData } : (formData || {});
  var tz = Session.getScriptTimeZone() || 'Asia/Hong_Kong';
  var results = [];

  // Helper: convert various timeCell formats into minutes from midnight
  function minutesFromTime(timeCell) {
    if (timeCell instanceof Date) return timeCell.getHours() * 60 + timeCell.getMinutes();
    if (typeof timeCell === 'number') {
      var frac = timeCell - Math.floor(timeCell);
      if (frac < 0) frac = 0;
      return Math.round(frac * 24 * 60);
    }
    if (typeof timeCell === 'string' && timeCell) {
      var s = timeCell.trim();
      var m24 = s.match(/^(\d{1,2}):(\d{2})$/);
      var m12 = s.match(/^(\d{1,2}):(\d{2})\s*(AM|PM)$/i);
      if (m24) return parseInt(m24[1], 10) * 60 + parseInt(m24[2], 10);
      if (m12) {
        var h = parseInt(m12[1], 10), m = parseInt(m12[2], 10);
        var ampm = m12[3].toUpperCase();
        if (ampm === 'PM' && h < 12) h += 12;
        if (ampm === 'AM' && h === 12) h = 0;
        return h * 60 + m;
      }
      var tmp = new Date(s);
      if (!isNaN(tmp.getTime())) return tmp.getHours() * 60 + tmp.getMinutes();
    }
    return 0;
  }

  // Helper: produce HH:MM string from timeCell
  function timeToHHMM(timeCell) {
    if (timeCell instanceof Date) {
      var th = timeCell.getHours(); var tm = timeCell.getMinutes();
      return (th < 10 ? '0' + th : th) + ':' + (tm < 10 ? '0' + tm : tm);
    }
    if (typeof timeCell === 'number') {
      var frac = timeCell - Math.floor(timeCell);
      if (frac < 0) frac = 0;
      var totalMinutes = Math.round(frac * 24 * 60);
      var nh = Math.floor(totalMinutes / 60);
      var nm = totalMinutes % 60;
      return (nh < 10 ? '0' + nh : nh) + ':' + (nm < 10 ? '0' + nm : nm);
    }
    if (typeof timeCell === 'string') {
      var s = timeCell.trim();
      var m24 = s.match(/^(\d{1,2}):(\d{2})$/);
      var m12 = s.match(/^(\d{1,2}):(\d{2})\s*(AM|PM)$/i);
      if (m24) {
        var hh = parseInt(m24[1], 10), mm = parseInt(m24[2], 10);
        return (hh < 10 ? '0' + hh : hh) + ':' + (mm < 10 ? '0' + mm : mm);
      }
      if (m12) {
        var hh2 = parseInt(m12[1], 10), mm2 = parseInt(m12[2], 10);
        var ap = m12[3].toUpperCase();
        if (ap === 'PM' && hh2 < 12) hh2 += 12;
        if (ap === 'AM' && hh2 === 12) hh2 = 0;
        return (hh2 < 10 ? '0' + hh2 : hh2) + ':' + (mm2 < 10 ? '0' + mm2 : mm2);
      }
      // fallback: try parse
      var tmp = new Date(s);
      if (!isNaN(tmp.getTime())) {
        var th = tmp.getHours(), tm = tmp.getMinutes();
        return (th < 10 ? '0' + th : th) + ':' + (tm < 10 ? '0' + tm : tm);
      }
      return s;
    }
    return '';
  }

  var fromDate = payload.from ? normalizeToDate(payload.from) : null;
  var toDate = payload.to ? normalizeToDate(payload.to) : null;
  if (!fromDate || !toDate) {
    // fallback to min/max in Days
    for (var k = 0; k < days.length; k++) {
      var dOnly = normalizeToDate(days[k].Date);
      if (!fromDate || dOnly.getTime() < fromDate.getTime()) fromDate = dOnly;
      if (!toDate || dOnly.getTime() > toDate.getTime()) toDate = dOnly;
    }
  }
  // Normalize times for inclusive compare
  if (fromDate) fromDate = new Date(fromDate.getFullYear(), fromDate.getMonth(), fromDate.getDate(), 0, 0, 0, 0);
  if (toDate) toDate = new Date(toDate.getFullYear(), toDate.getMonth(), toDate.getDate(), 23, 59, 59, 999);
  //Debug:
  console.log('Date Range:', { fromDate, toDate });
  console.log('Payload Items:', payload.items);

  (payload.items || []).forEach(function (item) {
    var day = (item.day || '').toString().trim();
    var periodId = (item.periodId || '').toString().trim();
    var title = (item.title || '').toString().trim();
    var location = (item.location || '').toString().trim();

    if (!day || !periodId || !title) return;

    // Find matching dates in Days
    var matchedDays = days.filter(function (d) {
      if (d.Day !== day) return false;
      var dateObj = normalizeToDate(d.Date);
      if (fromDate && dateObj < fromDate) return false;
      if (toDate && dateObj > toDate) return false;
      return true;
    });

    matchedDays.forEach(function (md) {
      var key = [md.DayType, periodId].join('|');
      var ts = dayTimesMap[key];
      if (!ts) return; // No time mapping

      var dateObj = normalizeToDate(md.Date); // ensure Date
      var minutesStart = minutesFromTime(ts.start);
      var minutesEnd = minutesFromTime(ts.end);
      var sortKey = dateObj.getTime() + minutesStart * 60000;

      var schoogleEventId = [md.CycleNum, md.Day, periodId, title].join('-');

      results.push({
        title: title,
        startDate: Utilities.formatDate(dateObj, tz, 'yyyy-MM-dd'),
        startTime: timeToHHMM(ts.start),
        endTime: timeToHHMM(ts.end),
        _sortKey: sortKey,
        location: location,
        schoogleEventId: schoogleEventId,
        day: md.Day,
        periodId: periodId,
        dayType: md.DayType,
        cycleNum: md.CycleNum,
        st: ts.start,
        et: ts.end
      });
    });

    
  });
    

  // Sort by computed sort key (date + start minutes)
  results.sort(function (a, b) { return (a._sortKey || 0) - (b._sortKey || 0); });
  // Remove internal keys and return serializable shape
  var serializableResults = results.map(function (r) {
    var copy = Object.assign({}, r);
    delete copy._sortKey;
    return copy;
  });
  console.log('SerializableResults:', serializableResults);
  return serializableResults;
}

function combineDateAndTime(dateOnly, timeCell, tz) {
  // timeCell could be a Date (time-of-day) or a string like "08:30" or "8:30 AM"
  var timeParts;
  if (timeCell instanceof Date) {
    var tH = timeCell.getHours();
    var tM = timeCell.getMinutes();
    return new Date(dateOnly.getFullYear(), dateOnly.getMonth(), dateOnly.getDate(), tH, tM, 0, 0);
  } else if (typeof timeCell === 'number') {
    // Google Sheets often returns times as a fractional day number (e.g. 0.3541667 -> 08:30)
    // Use only the fractional part to get time-of-day.
    var frac = timeCell - Math.floor(timeCell);
    if (frac < 0) frac = 0;
    var totalMinutes = Math.round(frac * 24 * 60);
    var nh = Math.floor(totalMinutes / 60);
    var nm = totalMinutes % 60;
    return new Date(dateOnly.getFullYear(), dateOnly.getMonth(), dateOnly.getDate(), nh, nm, 0, 0);
  } else if (typeof timeCell === 'string' && timeCell) {
    // Try to parse common formats
    var s = timeCell.trim();
    var m24 = s.match(/^(\d{1,2}):(\d{2})$/);
    var m12 = s.match(/^(\d{1,2}):(\d{2})\s*(AM|PM)$/i);
    var h = 0, m = 0;
    if (m24) {
      h = parseInt(m24[1], 10); m = parseInt(m24[2], 10);
    } else if (m12) {
      h = parseInt(m12[1], 10); m = parseInt(m12[2], 10);
      var ampm = m12[3].toUpperCase();
      if (ampm === 'PM' && h < 12) h += 12;
      if (ampm === 'AM' && h === 12) h = 0;
    } else {
      // Fallback: attempt Date parsing
      var tmp = new Date(s);
      if (!isNaN(tmp.getTime())) {
        h = tmp.getHours(); m = tmp.getMinutes();
      }
    }
    return new Date(dateOnly.getFullYear(), dateOnly.getMonth(), dateOnly.getDate(), h, m, 0, 0);
  }
  // Fallback to date start
  return new Date(dateOnly.getFullYear(), dateOnly.getMonth(), dateOnly.getDate(), 0, 0, 0, 0);
}

function normalizeToDate(value) {
  if (value instanceof Date) return value;
  if (typeof value === 'number') return new Date(Math.round(value));
  if (typeof value === 'string') {
    var d = new Date(value);
    if (!isNaN(d.getTime())) return d;
  }
  // default to today to avoid crashes
  return new Date();
}

function formatRFC3339(dt, tz) {
  return Utilities.formatDate(dt, tz, "yyyy-MM-dd'T'HH:mm:ssXXX");
}

function createCalendarEvents(eventsToCreate, calendarId) {
  var tz = Session.getScriptTimeZone() || 'Asia/Hong_Kong';
  var createdEventData = [];
  var errors = [];

  for (var i = 0; i < eventsToCreate.length; i++) {
    var ev = eventsToCreate[i];
    try {
      var resource = {
        summary: ev.title,
        location: ev.location || '',
        start: { dateTime: formatRFC3339(new Date(ev.startDateTime), tz), timeZone: tz },
        end: { dateTime: formatRFC3339(new Date(ev.endDateTime), tz), timeZone: tz },
        extendedProperties: { private: { Schoogle: 'true', SchoogleEventID: ev.schoogleEventId } }
      };
      var created = Calendar.Events.insert(resource, calendarId);
      createdEventData.push({
        Title: ev.title,
        Start: new Date(ev.startDateTime),
        End: new Date(ev.endDateTime),
        Location: ev.location || '',
        GoogleEventID: created.id,
        CalID: calendarId,
        Day: ev.day || '',
        Period: ev.periodId || '',
        DayType: ev.dayType || '',
        SchoogleEventID: ev.schoogleEventId
      });
    } catch (err) {
      errors.push({ index: i, message: err && err.message ? err.message : (err + '') });
    }
    Utilities.sleep(1000); // rate limit safety
  }
//Debug:
  console.log('Created Events:', createdEventData);
  console.log('Errors:', errors);

  // Report is now handled through the central Schoogle spreadsheet
  var reportUrl = null;
  // No need to create individual report spreadsheets - events are tracked in central spreadsheet

  return { success: errors.length === 0, createdEvents: createdEventData, errors: errors, reportUrl: reportUrl };
}

// --- Background job queue helpers ---------------------------------
// Legacy functions removed - now using central Schoogle spreadsheet approach

// --- Central Schoogle spreadsheet (per-user) helpers -------------------------
function getOrCreateSchoogleConfigFile_(folder) {
  var files = folder.getFilesByName('schoogle.json');
  if (files.hasNext()) return files.next();
  return folder.createFile('schoogle.json', '{}');
}

function getOrCreateSchoogleSpreadsheet() {
  var folder = getUserJobsFolder();
  var cfgFile = getOrCreateSchoogleConfigFile_(folder);
  var cfg = {};
  try { cfg = JSON.parse(cfgFile.getBlob().getDataAsString() || '{}'); } catch (e) { cfg = {}; }

  var ss, id = cfg.spreadsheetId || '';
  if (id) {
    try { ss = SpreadsheetApp.openById(id); } catch (e) { ss = null; }
  }
  if (!ss) {
    ss = SpreadsheetApp.create('Schoogle');
    // move into SchoogleJobs folder
    try {
      var sfile = DriveApp.getFileById(ss.getId());
      folder.addFile(sfile);
      DriveApp.getRootFolder().removeFile(sfile);
    } catch (e) { /* ignore */ }
    cfg.spreadsheetId = ss.getId();
    cfg.url = ss.getUrl();
    cfg.owner = Session.getActiveUser().getEmail() || Session.getEffectiveUser().getEmail() || '';
    cfg.updatedAt = new Date().toISOString();
    try { cfgFile.setContent(JSON.stringify(cfg)); } catch (e) { /* ignore */ }
  }

  // ensure JOBS sheet exists with header
  var sh = ss.getSheetByName('JOBS');
  if (!sh) sh = ss.insertSheet('JOBS');
  if (sh.getLastRow() === 0) {
    sh.getRange(1,1,1,18).setValues([[
      'jobId','status','calendarId','title','startDateTime','endDateTime','location',
      'day','periodId','dayType','schoogleEventId','createdEventId','error',
      'owner','createdAt','takenAt','doneAt','workerRunId'
    ]]);
  }
  // ensure Created Events sheet exists (central log of created events)
  var ce = ss.getSheetByName('Created Events');
  if (!ce) ce = ss.insertSheet('Created Events');
  if (ce.getLastRow() === 0) {
    ce.getRange(1,1,1,10).setValues([[
      'Title','Start','End','Location','GoogleEventID','CalID','Day','Period','DayType','SchoogleEventID'
    ]]);
  }
  // ensure Job Log exists (used to avoid duplicate completion emails)
  var jl = ss.getSheetByName('Job Log');
  if (!jl) jl = ss.insertSheet('Job Log');
  if (jl.getLastRow() === 0) {
    jl.getRange(1,1,1,5).setValues([[ 'jobId','timestamp','type','message','count' ]]);
  }
  return { id: ss.getId(), url: ss.getUrl(), ss: ss, folder: folder };
}

function getJobsSheetCentral_() {
  var sc = getOrCreateSchoogleSpreadsheet();
  var sh = sc.ss.getSheetByName('JOBS');
  if (!sh) sh = sc.ss.insertSheet('JOBS');
  if (sh.getLastRow() === 0) {
    sh.getRange(1,1,1,18).setValues([[
      'jobId','status','calendarId','title','startDateTime','endDateTime','location',
      'day','periodId','dayType','schoogleEventId','createdEventId','error',
      'owner','createdAt','takenAt','doneAt','workerRunId'
    ]]);
  }
  return { sh: sh, spreadsheetId: sc.id, spreadsheetUrl: sc.url };
}

// Column indices (1-based) for JOBS sheet
var JOBS_COL = {
  jobId: 1, status: 2, calendarId: 3, title: 4, start: 5, end: 6, location: 7,
  day: 8, period: 9, dayType: 10, schoogleEventId: 11, createdEventId: 12, error: 13,
  owner: 14, createdAt: 15, takenAt: 16, doneAt: 17, workerRunId: 18
};

function getSchoogleSpreadsheetUrl() {
  var sc = getOrCreateSchoogleSpreadsheet();
  return { id: sc.id, url: sc.url };
}

// Get detailed event-level job data for Active Jobs UI
function listUserJobEvents() {
  var out = [];
  try {
    var central = getJobsSheetCentral_();
    var sh = central.sh;
    var lastRow = sh.getLastRow();
    if (lastRow >= 2) {
      var vals = sh.getRange(2, 1, lastRow - 1, 18).getValues();
      vals.forEach(function(r, idx) {
        var jobId = r[JOBS_COL.jobId - 1] || '';
        var status = (r[JOBS_COL.status - 1] || '').toString();
        var calendarId = r[JOBS_COL.calendarId - 1] || 'primary';
        var title = r[JOBS_COL.title - 1] || '';
        var startDt = r[JOBS_COL.start - 1];
        var location = r[JOBS_COL.location - 1] || '';
        var createdAt = r[JOBS_COL.createdAt - 1] || '';
        var takenAt = r[JOBS_COL.takenAt - 1] || '';
        var doneAt = r[JOBS_COL.doneAt - 1] || '';
        var error = r[JOBS_COL.error - 1] || '';
        var createdEventId = r[JOBS_COL.createdEventId - 1] || '';
        
        if (jobId) {
          out.push({
            jobId: jobId,
            status: status,
            calendarId: calendarId,
            title: title,
            startDateTime: startDt instanceof Date ? startDt.toISOString() : startDt,
            location: location,
            createdAt: createdAt instanceof Date ? createdAt.toISOString() : createdAt,
            takenAt: takenAt instanceof Date ? takenAt.toISOString() : takenAt,
            doneAt: doneAt instanceof Date ? doneAt.toISOString() : doneAt,
            error: error,
            createdEventId: createdEventId // Include Google Calendar event ID for linking
          });
        }
      });
    }
  } catch (e) { 
    console.error('listUserJobEvents error:', e); 
    return []; // Return empty array instead of letting it continue
  }
  // Sort by creation time, newest first
  out.sort(function(a, b) { 
    return (new Date(b.createdAt || 0)) - (new Date(a.createdAt || 0)); 
  });
  return out;
}

// Append a single created event to the central Created Events sheet
function appendCreatedEventCentral_(ss, row) {
  try {
    var ce = ss.getSheetByName('Created Events');
    if (!ce) {
      ce = ss.insertSheet('Created Events');
      ce.getRange(1,1,1,10).setValues([[ 'Title','Start','End','Location','GoogleEventID','CalID','Day','Period','DayType','SchoogleEventID' ]]);
    }
    ce.appendRow([
      row.Title || '', row.Start || '', row.End || '', row.Location || '', row.GoogleEventID || '',
      row.CalID || '', row.Day || '', row.Period || '', row.DayType || '', row.SchoogleEventID || ''
    ]);
  } catch (e) { /* ignore */ }
}

// Build "Class Sessions" sheet by grouping Created Events dates under each Title
function rebuildClassSessionsFromCreated_(ss) {
  try {
    var ce = ss.getSheetByName('Created Events');
    if (!ce) return;
    var cs = ss.getSheetByName('Class Sessions');
    if (cs) ss.deleteSheet(cs);
    cs = ss.insertSheet('Class Sessions');
    var last = ce.getLastRow();
    if (last < 2) return;
    var vals = ce.getRange(2,1,last-1,10).getValues();
    var byTitle = {};
    for (var i=0;i<vals.length;i++) {
      var title = (vals[i][0] || '').toString();
      if (!title) continue;
      var start = vals[i][1];
      if (!byTitle[title]) byTitle[title] = [];
      byTitle[title].push(start instanceof Date ? start : new Date(start));
    }
    var titles = Object.keys(byTitle).sort();
    if (!titles.length) return;
    var maxRows = 0;
    titles.forEach(function(t){ byTitle[t].sort(function(a,b){ return a.getTime()-b.getTime(); }); if (byTitle[t].length>maxRows) maxRows = byTitle[t].length; });
    cs.getRange(1,1,1,titles.length).setValues([titles]);
    for (var r=0;r<maxRows;r++) {
      var row = [];
      for (var c=0;c<titles.length;c++) row.push(byTitle[titles[c]][r] || '');
      cs.getRange(2 + r, 1, 1, titles.length).setValues([row]);
    }
    try { if (maxRows>0) cs.getRange(2,1,maxRows,titles.length).setNumberFormat('dd-MMM (ddd) HH:mm'); } catch (e) {}
    cs.autoResizeColumns(1, titles.length);
  } catch (e) { /* ignore */ }
}

function jobCompletionAlreadyLogged_(ss, jobId) {
  try {
    var jl = ss.getSheetByName('Job Log');
    if (!jl) return false;
    var last = jl.getLastRow();
    if (last < 2) return false;
    var vals = jl.getRange(2,1,last-1,5).getValues();
    for (var i=0;i<vals.length;i++) {
      if ((vals[i][0]||'') === jobId && (vals[i][2]||'') === 'completed') return true;
    }
  } catch (e) { /* ignore */ }
  return false;
}

function logJobCompleted_(ss, jobId, count) {
  try {
    var jl = ss.getSheetByName('Job Log');
    if (!jl) {
      jl = ss.insertSheet('Job Log');
      jl.getRange(1,1,1,5).setValues([[ 'jobId','timestamp','type','message','count' ]]);
    }
    jl.appendRow([ jobId, new Date(), 'completed', 'Job completed', count || 0 ]);
  } catch (e) { /* ignore */ }
}

// Ensure a per-user worker trigger exists (creates an installable trigger under the
// calling user's account so jobs run with the user's credentials).
function ensureUserWorkerTrigger() {
  // Check existing project triggers and create an installable trigger for processUserJobs if missing.
  var projTriggers = ScriptApp.getProjectTriggers();
  for (var i = 0; i < projTriggers.length; i++) {
    try {
      if (projTriggers[i].getHandlerFunction && projTriggers[i].getHandlerFunction() === 'processUserJobs') return;
    } catch (e) { /* ignore */ }
  }
  // create a minute-based trigger running as the current user
  // run less frequently to reduce polling when idle
  ScriptApp.newTrigger('processUserJobs').timeBased().everyMinutes(1).create();
}

function enqueueJob(events, calendarId) {
  // Centralized: append events to per-user Schoogle spreadsheet (JOBS sheet)
  var jobId = 'job_' + Utilities.getUuid();
  var owner = Session.getActiveUser().getEmail() || Session.getEffectiveUser().getEmail() || '';
  var now = new Date();
  var central = getJobsSheetCentral_();
  var sh = central.sh;

  if (!Array.isArray(events) || events.length === 0) return jobId;

  var rows = events.map(function(ev) {
    return [
      jobId, 'queued', (calendarId || 'primary'),
      (ev.title || ''),
      new Date(ev.startDateTime), new Date(ev.endDateTime),
      (ev.location || ''),
      (ev.day || ''), (ev.periodId || ''), (ev.dayType || ''),
      (ev.schoogleEventId || ''),
      '', // createdEventId
      '', // error
      owner,
      now, // createdAt
      '',  // takenAt
      '',  // doneAt
      ''   // workerRunId
    ];
  });

  var lock = LockService.getScriptLock();
  try {
    lock.waitLock(30000);
    var startRow = sh.getLastRow() + 1;
    sh.getRange(startRow, 1, rows.length, rows[0].length).setValues(rows);
  } finally {
    try { lock.releaseLock(); } catch (e) {}
  }

  try { ensureUserWorkerTrigger(); } catch (e) { /* ignore */ }
  return jobId;
}

// Remove the per-user installable trigger for processUserJobs if there are no queued/running jobs
function _removeUserWorkerTriggerIfIdle() {
  try {
    var central = getJobsSheetCentral_();
    var sh = central.sh;
    var lastRow = sh.getLastRow();
    var hasActive = false;
    if (lastRow >= 2) {
      var vals = sh.getRange(2, JOBS_COL.status, lastRow - 1, 1).getValues();
      for (var i = 0; i < vals.length; i++) {
        var st = (vals[i][0] || '').toString();
        if (st === 'queued' || st === 'running') { hasActive = true; break; }
      }
    }
    if (!hasActive) {
      var projTriggers = ScriptApp.getProjectTriggers();
      for (var t = 0; t < projTriggers.length; t++) {
        if (projTriggers[t].getHandlerFunction && projTriggers[t].getHandlerFunction() === 'processUserJobs') {
          ScriptApp.deleteTrigger(projTriggers[t]);
        }
      }
    }
  } catch (e) { /* ignore */ }
}

// --- Legacy job spreadsheet helpers removed ------------------------------------
// Functions removed: createJobSpreadsheetForJob, appendJobLog, appendCreatedEvents
// Now using central Schoogle spreadsheet approach with appendCreatedEventCentral_

function readMetaFlag(spreadsheetId, key) {
  try {
    var ss = SpreadsheetApp.openById(spreadsheetId);
    var meta = ss.getSheetByName('Meta');
    if (!meta) return null;
    var vals = meta.getRange(1,1,meta.getLastRow(),2).getValues();
    for (var i=0;i<vals.length;i++){
      if ((vals[i][0]||'') === key) return vals[i][1];
    }
  } catch (e) { return null; }
  return null;
}

function setMetaFlag(spreadsheetId, key, value) {
  try {
    var ss = SpreadsheetApp.openById(spreadsheetId);
    var meta = ss.getSheetByName('Meta');
    if (!meta) meta = ss.insertSheet('Meta');
    var vals = meta.getRange(1,1,meta.getLastRow(),2).getValues();
    var found = false;
    for (var i=0;i<vals.length;i++){
      if ((vals[i][0]||'') === key) { meta.getRange(i+1,2).setValue(value); found = true; break; }
    }
    if (!found) meta.getRange(meta.getLastRow()+1,1,1,2).setValues([[key,value]]);
  } catch (e) { /* ignore */ }
}

function writeReportToSpreadsheet(spreadsheetId, reportData) {
  try {
    // overwrite Created Events sheet with reportData
    var ss = SpreadsheetApp.openById(spreadsheetId);
    var sheet = ss.getSheetByName('Created Events');
    if (!sheet) sheet = ss.insertSheet('Created Events');
    // clear except header
    var last = sheet.getLastRow();
    if (last > 1) sheet.getRange(2,1,last-1,sheet.getLastColumn()).clearContent();
    if (reportData && reportData.length) {
      var values = reportData.map(function (r) { return [r.Title, r.Start, r.End, r.Location, r.GoogleEventID, r.CalID, r.Day, r.Period, r.DayType, r.SchoogleEventID]; });
      sheet.getRange(2,1,values.length,values[0].length).setValues(values);
    }
    // also set Meta.reportUrl if applicable
    try { setMetaFlag(spreadsheetId, 'reportUrl', ss.getUrl()); } catch (e) { /* ignore */ }
    // Also create a "Class Sessions" sheet grouping events by Title into columns
    try {
      var cs = ss.getSheetByName('Class Sessions');
      if (cs) {
        // remove existing data
        ss.deleteSheet(cs);
      }
      cs = ss.insertSheet('Class Sessions');
      // Collect unique titles
      var byTitle = {};
      (reportData || []).forEach(function (r) {
        var t = (r.Title || '').toString();
        if (!t) return;
        byTitle[t] = byTitle[t] || [];
        byTitle[t].push(new Date(r.Start));
      });
      var titles = Object.keys(byTitle).sort();
      var maxRows = 0;
      titles.forEach(function (t) {
        byTitle[t].sort(function (a, b) { return a.getTime() - b.getTime(); });
        if (byTitle[t].length > maxRows) maxRows = byTitle[t].length;
      });
      if (titles.length) {
        // Headers
        cs.getRange(1, 1, 1, titles.length).setValues([titles]);
        // Fill columns as rows under headers
        for (var r = 0; r < maxRows; r++) {
          var rowVals = [];
          for (var c = 0; c < titles.length; c++) {
            rowVals.push(byTitle[titles[c]][r] || '');
          }
          cs.getRange(2 + r, 1, 1, titles.length).setValues([rowVals]);
        }
        cs.autoResizeColumns(1, titles.length);
        // Format date/time cells from row 2 onwards
        try {
          if (maxRows > 0) cs.getRange(2, 1, maxRows, titles.length).setNumberFormat('dd-MMM (ddd) HH:mm');
        } catch (e) { /* ignore */ }
      }
    } catch (e) { /* ignore non-fatal sheet generation errors */ }

    return ss.getUrl();
  } catch (e) { return null; }
}

// Legacy function: getJobStatus - deprecated in favor of listUserJobs()
// Kept for backwards compatibility but client now uses listUserJobs() exclusively
function getJobStatus(jobId) {
  // Delegate to listUserJobs and filter for the specific job
  var allJobs = listUserJobs();
  var matchingJob = allJobs.find(function(j) { return j.jobId === jobId; });
  if (!matchingJob) return { jobId: jobId, status: 'notfound' };
  
  return {
    jobId: jobId,
    status: matchingJob.status,
    progress: { done: matchingJob.done, total: matchingJob.total, created: matchingJob.done, errors: matchingJob.errors, estSeconds: null },
    reportUrl: matchingJob.reportUrl,
    spreadsheetUrl: matchingJob.spreadsheetUrl,
    errors: matchingJob.errors
  };
}

// List all jobs for the current user (metadata only)
function listUserJobs() {
  var out = [];
  try {
    var central = getJobsSheetCentral_();
    var sh = central.sh;
    var lastRow = sh.getLastRow();
    if (lastRow >= 2) {
      var vals = sh.getRange(2, 1, lastRow - 1, 18).getValues();
      var byJob = {};
      vals.forEach(function(r){
        var jid = r[JOBS_COL.jobId - 1];
        if (!jid) return;
        var st = (r[JOBS_COL.status - 1] || '').toString();
        var createdAt = r[JOBS_COL.createdAt - 1] || '';
        if (!byJob[jid]) byJob[jid] = { jobId: jid, total: 0, done: 0, errors: 0, lastUpdated: '', status: 'queued', spreadsheetUrl: central.spreadsheetUrl, reportUrl: central.spreadsheetUrl, startedAt: '' };
        byJob[jid].total += 1;
        if (st === 'done') byJob[jid].done += 1;
        if (st === 'error') byJob[jid].errors += 1;
        if (st === 'running') byJob[jid].status = 'running';
        var ts = r[JOBS_COL.doneAt - 1] || r[JOBS_COL.takenAt - 1] || createdAt || '';
        if (ts) byJob[jid].lastUpdated = ts instanceof Date ? ts.toISOString() : ts;
        var startedAt = r[JOBS_COL.takenAt - 1] || createdAt || '';
        if (st === 'running' && !byJob[jid].startedAt) byJob[jid].startedAt = startedAt instanceof Date ? startedAt.toISOString() : startedAt;
      });
      Object.keys(byJob).forEach(function(k){
        var j = byJob[k];
        if (j.done + j.errors === j.total && j.status !== 'running') j.status = 'done';
        out.push({ 
          jobId: j.jobId, 
          status: j.status, 
          done: j.done, 
          total: j.total, 
          reportUrl: j.reportUrl, 
          spreadsheetUrl: j.spreadsheetUrl, 
          createdCount: j.done, 
          errors: j.errors, 
          startedAt: j.startedAt, 
          lastUpdated: j.lastUpdated 
        });
      });
    }
  } catch (e) { /* ignore */ }
  out.sort(function(a,b){ return (new Date(b.lastUpdated || 0)) - (new Date(a.lastUpdated || 0)); });

  // cleanup triggers if no active
  if (!out.some(function(j){ return j.status === 'queued' || j.status === 'running'; })) {
    try { _removeUserWorkerTriggerIfIdle(); } catch (e) { /* ignore */ }
  }
  return out;
}

// Helper: get or create a SchoogleJobs folder in the current user's My Drive
function getUserJobsFolder() {
  var FOLDER_NAME = 'SchoogleJobs';
  var folders = DriveApp.getFoldersByName(FOLDER_NAME);
  if (folders.hasNext()) return folders.next();
  return DriveApp.createFolder(FOLDER_NAME);
}

// Worker to be installed per-user: scans the user's SchoogleJobs folder and processes queued jobs.
function processUserJobs() {
  //log into the log who is running the process
  try {Logger.log(Session.getActiveUser().getEmail())} catch (e) { Logger.log(e)/* ignore */ }
  // Prevent concurrent runs per user
  var userLock = LockService.getUserLock();
  if (!userLock.tryLock(30000)) return;
  try {
    var central = getJobsSheetCentral_();
    var sh = central.sh;
    var tz = Session.getScriptTimeZone() || 'Asia/Hong_Kong';
    var CHUNK = 12;

    var lastRow = sh.getLastRow();
    if (lastRow < 2) return; // nothing queued

    // Claim next CHUNK queued rows under a script-level lock
    var scriptLock = LockService.getScriptLock();
    scriptLock.waitLock(30000);
  var vals = sh.getRange(2, 1, lastRow - 1, 18).getValues();
  // Use a timestamp-based run id for traceability, e.g. 20250820_143512_123
  var workerRunId = Utilities.formatDate(new Date(), tz, 'yyyyMMdd_HHmmss_SSS');
  var now = new Date();
    var claimIndexes = [];
    for (var i = 0; i < vals.length && claimIndexes.length < CHUNK; i++) {
      if ((vals[i][JOBS_COL.status - 1] || '').toString() === 'queued') {
        claimIndexes.push(i);
        vals[i][JOBS_COL.status - 1] = 'running';
        vals[i][JOBS_COL.takenAt - 1] = now;
        vals[i][JOBS_COL.workerRunId - 1] = workerRunId;
      }
    }
    if (claimIndexes.length) {
      // Write back claims (set non-contiguous columns individually)
      for (var ci = 0; ci < claimIndexes.length; ci++) {
        var r = claimIndexes[ci];
        var sheetRow = r + 2;
        // status -> running
        sh.getRange(sheetRow, JOBS_COL.status).setValue('running');
        // takenAt -> now
        sh.getRange(sheetRow, JOBS_COL.takenAt).setValue(now);
        // workerRunId -> id
        sh.getRange(sheetRow, JOBS_COL.workerRunId).setValue(workerRunId);
      }
    }
    scriptLock.releaseLock();

    if (!claimIndexes.length) {
      // no queued rows; consider removing trigger
      try { _removeUserWorkerTriggerIfIdle(); } catch (e) { /* ignore */ }
      return;
    }

    // Process claimed rows
    var sc = getOrCreateSchoogleSpreadsheet();
    var processedJobIds = {};
    for (var k = 0; k < claimIndexes.length; k++) {
      var idx = claimIndexes[k];
      var rowNum = idx + 2; // sheet row number
      var row = vals[idx];
      var jobId = row[JOBS_COL.jobId - 1] || '';
      var calendarId = row[JOBS_COL.calendarId - 1] || 'primary';
      var title = row[JOBS_COL.title - 1] || '';
      var location = row[JOBS_COL.location - 1] || '';
      var schoogleEventId = row[JOBS_COL.schoogleEventId - 1] || '';
      var startDt = new Date(row[JOBS_COL.start - 1]);
      var endDt = new Date(row[JOBS_COL.end - 1]);
      var createdEventId = '';
      var errorMsg = '';

      try {
        var resource = {
          summary: title,
          location: location,
          start: { dateTime: formatRFC3339(startDt, tz), timeZone: tz },
          end: { dateTime: formatRFC3339(endDt, tz), timeZone: tz },
          extendedProperties: { private: { Schoogle: 'true', SchoogleEventID: schoogleEventId } }
        };
        var created = Calendar.Events.insert(resource, calendarId);
        createdEventId = created.id || '';
      } catch (err) {
        errorMsg = (err && err.message) ? err.message : String(err);
      }
      Utilities.sleep(1000);

  // Update status cells for this row (set non-contiguous columns correctly)
  var status = createdEventId ? 'done' : 'error';
  var doneAt = new Date();
  // createdEventId
  sh.getRange(rowNum, JOBS_COL.createdEventId).setValue(createdEventId);
  // error
  sh.getRange(rowNum, JOBS_COL.error).setValue(errorMsg);
  // doneAt
  sh.getRange(rowNum, JOBS_COL.doneAt).setValue(doneAt);
  // workerRunId (keep for traceability)
  sh.getRange(rowNum, JOBS_COL.workerRunId).setValue(workerRunId);
  // status
  sh.getRange(rowNum, JOBS_COL.status).setValue(status);
      // If event was created, append to central Created Events
      if (createdEventId) {
        appendCreatedEventCentral_(sc.ss, {
          Title: title,
          Start: startDt,
          End: endDt,
          Location: location,
          GoogleEventID: createdEventId,
          CalID: calendarId,
          Day: row[JOBS_COL.day - 1] || '',
          Period: row[JOBS_COL.period - 1] || '',
          DayType: row[JOBS_COL.dayType - 1] || '',
          SchoogleEventID: schoogleEventId
        });
      }
      if (jobId) processedJobIds[jobId] = true;
    }

    // Try to stop trigger if idle
    try { _removeUserWorkerTriggerIfIdle(); } catch (e) { /* ignore */ }

    // For each processed job, if it's now complete, rebuild Class Sessions and send one completion email
    try {
      var allLast = sh.getLastRow();
      var allVals = allLast > 1 ? sh.getRange(2,1,allLast-1,18).getValues() : [];
      var ownerEmail = Session.getActiveUser().getEmail() || Session.getEffectiveUser().getEmail() || '';
      for (var jid in processedJobIds) {
        var anyActive = false, total = 0, createdCount = 0, errorCount = 0;
        for (var i2=0;i2<allVals.length;i2++) {
          if ((allVals[i2][JOBS_COL.jobId - 1] || '') !== jid) continue;
          total++;
          var st2 = (allVals[i2][JOBS_COL.status - 1] || '').toString();
          if (st2 === 'queued' || st2 === 'running') anyActive = true;
          if (st2 === 'done' && (allVals[i2][JOBS_COL.createdEventId - 1] || '')) createdCount++;
          if (st2 === 'error') errorCount++;
        }
        if (!anyActive && total > 0) {
          // All rows for this job are finished
          rebuildClassSessionsFromCreated_(sc.ss);
          if (!jobCompletionAlreadyLogged_(sc.ss, jid)) {
            try {
              if (ownerEmail) {
                var subj = 'Schoogle: Events Created!';
                var body = 'Your Schoogle job has completed.\n\nCreated events: ' + createdCount + '\nErrors: ' + errorCount + '\nSpreadsheet: ' + sc.url + '\n';
                MailApp.sendEmail(ownerEmail, subj, body);
              }
            } catch (mailErr) { /* ignore */ }
            logJobCompleted_(sc.ss, jid, createdCount);
          }
        }
      }
    } catch (postErr) { /* ignore */ }
  } finally {
    try { userLock.releaseLock(); } catch (e) {}
  }
}

// Worker: run under a time-based trigger. Processes small chunks and persists progress.
// Legacy function processJobs() removed - now using processUserJobs() with central spreadsheet

// Legacy function removed: createReportSpreadsheet
// Now using central Schoogle spreadsheet with Created Events and Class Sessions sheets

function getEventsFromCalendar(calendarId) {
  var out = [];
  var pageToken;
  try {
  do {
    var resp = Calendar.Events.list(calendarId, {
      privateExtendedProperty: 'Schoogle=true',
      maxResults: 2500,
      pageToken: pageToken,
      singleEvents: true,
      orderBy: 'startTime'
    });
    var items = (resp && resp.items) || [];
    items.forEach(function (ev) {
      var startDT = ev.start && (ev.start.dateTime || ev.start.date);
      var endDT = ev.end && (ev.end.dateTime || ev.end.date);
      out.push({
        id: ev.id,
        title: ev.summary || '',
        start: startDT,
        end: endDT,
        location: ev.location || '',
        htmlLink: ev.htmlLink || '',
        schoogle: ev.extendedProperties && ev.extendedProperties.private && ev.extendedProperties.private.Schoogle === 'true',
        schoogleEventId: ev.extendedProperties && ev.extendedProperties.private && ev.extendedProperties.private.SchoogleEventID
      });
    });
    pageToken = resp.nextPageToken;
  } while (pageToken);
  } catch (e) {
    // If calendar not found or access denied, return empty list instead of throwing
    try {
      var msg = e && e.message ? e.message.toString() : (e + '');
      if (msg && (msg.indexOf('Not Found') !== -1 || msg.indexOf('Not Found') !== -1 || msg.indexOf('notFound') !== -1)) {
        return out; // empty
      }
    } catch (e2) { /* ignore */ }
    return out;
  }

  return out;
}

function updateCalendarEvent(calendarId, eventId, newTitle) {
  var updated = Calendar.Events.patch({ summary: newTitle }, calendarId, eventId);
  return { success: true, id: updated.id, title: updated.summary };
}

function deleteCalendarEvent(calendarId, eventId) {
  Calendar.Events.remove(calendarId, eventId);
  return { success: true };
}

// Delete all Schoogle-created events in a calendar that match the provided title.
function deleteEventsByTitle(calendarId, title) {
  var out = { deleted: 0, errors: [] };
  if (!calendarId) calendarId = 'primary';
  var pageToken = null;
  do {
    var resp = Calendar.Events.list(calendarId, { privateExtendedProperty: 'Schoogle=true', q: title, maxResults: 2500, pageToken: pageToken });
    var items = (resp && resp.items) || [];
    items.forEach(function(ev) {
      try {
        // match exact title or contains
        if (!title || (ev.summary && (ev.summary === title || ev.summary.indexOf(title) !== -1))) {
          Calendar.Events.remove(calendarId, ev.id);
          out.deleted += 1;
        }
      } catch (e) {
        out.errors.push({ id: ev.id, message: e && e.message ? e.message : (e + '') });
      }
    });
    pageToken = resp && resp.nextPageToken;
  } while (pageToken);
  return out;
}

// Return array of Schoogle events matching title with basic metadata (id, title, start, end, location)
function listSchoogleEventsByTitle(calendarId, title) {
  var out = [];
  if (!calendarId) calendarId = 'primary';
  var pageToken = null;
  do {
    var resp = Calendar.Events.list(calendarId, { privateExtendedProperty: 'Schoogle=true', q: title, maxResults: 2500, pageToken: pageToken });
    var items = (resp && resp.items) || [];
    items.forEach(function(ev) {
      try {
        out.push({ id: ev.id, title: ev.summary || '', start: (ev.start && (ev.start.dateTime || ev.start.date)) || '', end: (ev.end && (ev.end.dateTime || ev.end.date)) || '', location: ev.location || '' });
      } catch (e) { /* ignore */ }
    });
    pageToken = resp && resp.nextPageToken;
  } while (pageToken);
  return out;
}

// Delete multiple event ids in the given calendar
function deleteEventsByIds(calendarId, ids) {
  var out = { deleted: 0, errors: [] };
  if (!calendarId) calendarId = 'primary';
  (ids || []).forEach(function(id) {
    try {
      Calendar.Events.remove(calendarId, id);
      out.deleted += 1;
    } catch (e) {
      out.errors.push({ id: id, message: e && e.message ? e.message : (e + '') });
    }
  });
  return out;
}
