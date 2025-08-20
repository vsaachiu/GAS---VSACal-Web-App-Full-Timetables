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
  if (ttSheet && email) {
    var tVals = ttSheet.getDataRange().getValues();
    for (var r = 1; r < tVals.length; r++) {
      var row = tVals[r];
      if ((row[0] || '').toString().trim().toLowerCase() === email.toLowerCase()) {
        userTimetable.push({
          day: (row[4] || '').toString().trim(),
          periodId: (row[2] || '').toString().trim(),
          title: (row[3] || '').toString().trim(),
          location: (row[6] || '').toString().trim()
        });
      }
    }
  }

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

  return { periods: periods, userTimetable: userTimetable, userEmail: email, defaultFrom: defaultFrom, defaultTo: defaultTo };
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

  // Generate report spreadsheet on the server so it is created even if the client disconnects
  var reportUrl = null;
  try {
    reportUrl = createReportSpreadsheet(createdEventData);
    console.log('Report URL:', reportUrl);
  } catch (repErr) {
    console.error('createReportSpreadsheet failed', repErr);
  }

  return { success: errors.length === 0, createdEvents: createdEventData, errors: errors, reportUrl: reportUrl };
}

// --- Background job queue helpers ---------------------------------
function getJobsSheet() {
  var PROP = PropertiesService.getScriptProperties().getProperty('JOBS_SPREADSHEET_ID');
  var ss;
  try {
    if (PROP) ss = SpreadsheetApp.openById(PROP);
  } catch (e) { ss = null; }
  if (!ss) ss = SpreadsheetApp.getActive();
  var name = 'Schoogle_Jobs';
  var sh = ss.getSheetByName(name);
  if (!sh) {
    sh = ss.insertSheet(name);
    var headers = ['jobId','payloadFileId','createdAt','status','done','total','reportUrl','errorsFileId','createdEventsFileId','owner'];
    sh.getRange(1, 1, 1, headers.length).setValues([headers]);
  }
  return sh;
}

function setJobsSpreadsheetId(id) {
  if (!id) throw new Error('Invalid id');
  PropertiesService.getScriptProperties().setProperty('JOBS_SPREADSHEET_ID', id);
  return id;
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
  ScriptApp.newTrigger('processUserJobs').timeBased().everyMinutes(1).create();
}

function enqueueJob(events, calendarId) {
  // Per-user enqueue: store job in the user's Drive under a SchoogleJobs folder and
  // create a per-user trigger to process jobs. This ensures Calendar inserts run
  // with the user's credentials (so events appear in their calendar).
  var jobId = 'job_' + Utilities.getUuid();
  var owner = Session.getActiveUser().getEmail() || Session.getEffectiveUser().getEmail();
  var payload = { jobId: jobId, events: events || [], calendarId: calendarId || 'primary', owner: owner, createdAt: (new Date()).toISOString(), status: 'queued', done: 0, total: (events || []).length, reportUrl: null, errors: [], createdEvents: [], startedAt: null, lastUpdated: (new Date()).toISOString() };

  // Ensure user's SchoogleJobs folder exists
  var folder = getUserJobsFolder();
  var file = folder.createFile(jobId + '.json', JSON.stringify(payload));

  // Create per-user trigger if not present
  try { ensureUserWorkerTrigger(); } catch (e) { /* non-fatal */ }

  return jobId;
}

function getJobStatus(jobId) {
  // Per-user job status: look for job file in the user's SchoogleJobs folder
  try {
    var folder = getUserJobsFolder();
    var files = folder.getFilesByName(jobId + '.json');
    if (!files.hasNext()) return { jobId: jobId, status: 'notfound' };
    var file = files.next();
    var payload = JSON.parse(file.getBlob().getDataAsString());
    var status = payload.status || 'unknown';
    var done = Number(payload.done) || 0;
    var total = Number(payload.total) || 0;
    var reportUrl = payload.reportUrl || null;
    var errors = payload.errors || [];
    var createdCount = (payload.createdEvents && payload.createdEvents.length) || 0;
    // Estimate remaining time (seconds) based on startedAt if available
    var estSeconds = null;
    try {
      if (payload.startedAt && done > 0) {
        var started = new Date(payload.startedAt).getTime();
        var now = new Date().getTime();
        var elapsed = Math.max(1, Math.round((now - started) / 1000));
        var rate = elapsed / done; // seconds per item
        estSeconds = Math.round(rate * Math.max(0, total - done));
      }
    } catch (e) { estSeconds = null; }
    return { jobId: jobId, status: status, progress: { done: done, total: total, created: createdCount, estSeconds: estSeconds }, reportUrl: reportUrl, errors: errors };
  } catch (e) {
    return { jobId: jobId, status: 'error', error: e.toString() };
  }
}

// List all jobs for the current user (metadata only)
function listUserJobs() {
  var out = [];
  try {
    var folder = getUserJobsFolder();
    var files = folder.getFiles();
    while (files.hasNext()) {
      var f = files.next();
      if (!f.getName().match(/^job_/)) continue;
      try {
        var payload = JSON.parse(f.getBlob().getDataAsString());
      } catch (e) { continue; }
      out.push({ jobId: payload.jobId || f.getName().replace(/\.json$/, ''), status: payload.status || 'unknown', done: Number(payload.done) || 0, total: Number(payload.total) || 0, reportUrl: payload.reportUrl || null, createdCount: (payload.createdEvents && payload.createdEvents.length) || 0, errors: (payload.errors && payload.errors.length) || 0, startedAt: payload.startedAt || null, lastUpdated: payload.lastUpdated || null });
    }
  } catch (e) { /* ignore */ }
  // sort by createdAt desc if present
  out.sort(function(a,b){ return (b.lastUpdated || '') .localeCompare(a.lastUpdated || ''); });
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
  var lock = LockService.getUserLock();
  if (!lock.tryLock(30000)) return; // already running for this user
  try {
    var folder = getUserJobsFolder();
    var files = folder.getFiles();
    var tz = Session.getScriptTimeZone() || 'Asia/Hong_Kong';
    var CHUNK = 5;
    while (files.hasNext()) {
      var f = files.next();
      if (!f.getName().match(/^job_/)) continue;
      try {
        var payload = JSON.parse(f.getBlob().getDataAsString());
      } catch (e) { continue; }
      if (!payload || payload.status !== 'queued' && payload.status !== 'running') continue;
  var events = payload.events || [];
      var total = events.length;
  payload.status = 'running';
  if (!payload.done) payload.done = 0;
  if (!payload.startedAt) payload.startedAt = (new Date()).toISOString();
      // process a chunk
      for (var idx = payload.done; idx < Math.min(payload.done + CHUNK, total); idx++) {
        var ev = events[idx];
        try {
          var resource = {
            summary: ev.title,
            location: ev.location || '',
            start: { dateTime: formatRFC3339(new Date(ev.startDateTime), tz), timeZone: tz },
            end: { dateTime: formatRFC3339(new Date(ev.endDateTime), tz), timeZone: tz },
            description: 'SchoogleJob:' + (payload.jobId || '') + ':' + idx,
            extendedProperties: { private: { Schoogle: 'true', SchoogleEventID: ev.schoogleEventId } }
          };
          var created = Calendar.Events.insert(resource, payload.calendarId || 'primary');
          payload.createdEvents = payload.createdEvents || [];
          payload.createdEvents.push({ Title: ev.title, Start: new Date(ev.startDateTime), End: new Date(ev.endDateTime), Location: ev.location || '', GoogleEventID: created.id, CalID: payload.calendarId || 'primary', Day: ev.day || '', Period: ev.periodId || '', DayType: ev.dayType || '', SchoogleEventID: ev.schoogleEventId });
        } catch (err) {
          payload.errors = payload.errors || [];
          payload.errors.push({ index: idx, message: err && err.message ? err.message : (err + '') });
        }
  Utilities.sleep(1000);
  payload.done = idx + 1;
  payload.lastUpdated = (new Date()).toISOString();
  // persist updated payload back to file
  try { f.setContent(JSON.stringify(payload)); } catch (e) { /* ignore */ }
      }

      // finalize if complete
      if (payload.done >= total) {
        try {
          var reportUrl = createReportSpreadsheet(payload.createdEvents || []);
          payload.reportUrl = reportUrl;
          payload.status = 'done';
          f.setContent(JSON.stringify(payload));
          // notify user
          try {
            var ownerEmail = Session.getActiveUser().getEmail() || Session.getEffectiveUser().getEmail();
            if (ownerEmail) {
              var subject = 'Schoogle: Job ' + (payload.jobId || '') + ' completed';
              var body = 'Your Schoogle job has completed.\n\nCreated events: ' + ((payload.createdEvents && payload.createdEvents.length) || 0) + '\nReport: ' + (reportUrl || '') + '\n\nIf you did not expect this, contact the administrator.';
              MailApp.sendEmail(ownerEmail, subject, body);
            }
          } catch (mailErr) { /* ignore */ }
        } catch (finalErr) {
          payload.status = 'error';
          payload.errors = payload.errors || [];
          payload.errors.push({ finalization: finalErr.toString() });
          try { f.setContent(JSON.stringify(payload)); } catch (e) { /* ignore */ }
          try {
            var ownerEmail2 = Session.getActiveUser().getEmail() || Session.getEffectiveUser().getEmail();
            if (ownerEmail2) {
              var subject2 = 'Schoogle: Job ' + (payload.jobId || '') + ' failed';
              var body2 = 'Your Schoogle job encountered an error during finalization.\n\nError: ' + finalErr.toString() + '\nPlease check the job file in your Drive.';
              MailApp.sendEmail(ownerEmail2, subject2, body2);
            }
          } catch (mailErr2) { /* ignore */ }
        }
      }
    }
  } finally {
    try { lock.releaseLock(); } catch (e) { /* ignore */ }
  }
}

// Worker: run under a time-based trigger. Processes small chunks and persists progress.
function processJobs() {
  var lock = LockService.getScriptLock();
  if (!lock.tryLock(30000)) return; // already running elsewhere
  try {
    var sh = getJobsSheet();
    var vals = sh.getDataRange().getValues();
    var pickRow = -1;
    for (var i = 1; i < vals.length; i++) {
      if ((vals[i][3] || '').toString() === 'queued') { pickRow = i + 1; break; }
    }
    if (pickRow === -1) return; // nothing to do

    // mark running
    sh.getRange(pickRow, 4).setValue('running');

    var payloadFileId = sh.getRange(pickRow, 2).getValue();
    var payload = {};
    try { payload = JSON.parse(DriveApp.getFileById(payloadFileId).getBlob().getDataAsString()); } catch (e) { payload = {}; }
    var events = payload.events || [];
    var total = events.length;
    var done = Number(sh.getRange(pickRow, 5).getValue()) || 0;

    var createdFileId = sh.getRange(pickRow, 9).getValue() || '';
    var createdArr = [];
    if (createdFileId) {
      try { createdArr = JSON.parse(DriveApp.getFileById(createdFileId).getBlob().getDataAsString()) || []; } catch (e) { createdArr = []; }
    }

    var errorsFileId = sh.getRange(pickRow, 8).getValue() || '';
    var errorsArr = [];
    if (errorsFileId) {
      try { errorsArr = JSON.parse(DriveApp.getFileById(errorsFileId).getBlob().getDataAsString()) || []; } catch (e) { errorsArr = []; }
    }

    var CHUNK = 5;
    var tz = Session.getScriptTimeZone() || 'Asia/Hong_Kong';
    for (var idx = done; idx < Math.min(done + CHUNK, total); idx++) {
      var ev = events[idx];
      try {
        var resource = {
          summary: ev.title,
          location: ev.location || '',
          start: { dateTime: formatRFC3339(new Date(ev.startDateTime), tz), timeZone: tz },
          end: { dateTime: formatRFC3339(new Date(ev.endDateTime), tz), timeZone: tz },
          description: 'SchoogleJob:' + (sh.getRange(pickRow,1).getValue() || '') + ':' + idx,
          extendedProperties: { private: { Schoogle: 'true', SchoogleEventID: ev.schoogleEventId } }
        };
        var created = Calendar.Events.insert(resource, payload.calendarId || 'primary');
        createdArr.push({ Title: ev.title, Start: new Date(ev.startDateTime), End: new Date(ev.endDateTime), Location: ev.location || '', GoogleEventID: created.id, CalID: payload.calendarId || 'primary', Day: ev.day || '', Period: ev.periodId || '', DayType: ev.dayType || '', SchoogleEventID: ev.schoogleEventId });
      } catch (err) {
        errorsArr.push({ index: idx, message: err && err.message ? err.message : (err + '') });
      }
      Utilities.sleep(1000);
      // persist progress after each item
      sh.getRange(pickRow, 5).setValue(idx + 1);
    }

    // persist createdArr
    if (createdArr.length) {
      if (createdFileId) {
        try {
          var existing = JSON.parse(DriveApp.getFileById(createdFileId).getBlob().getDataAsString()) || [];
        } catch (e) { existing = []; }
        var merged = existing.concat(createdArr);
        DriveApp.getFileById(createdFileId).setContent(JSON.stringify(merged));
      } else {
        var f = DriveApp.createFile((sh.getRange(pickRow,1).getValue() || 'job') + '_created.json', JSON.stringify(createdArr));
        sh.getRange(pickRow, 9).setValue(f.getId());
      }
    }

    // persist errors
    if (errorsArr.length) {
      if (errorsFileId) {
        try { var existingE = JSON.parse(DriveApp.getFileById(errorsFileId).getBlob().getDataAsString()) || []; } catch (e) { existingE = []; }
        var mergedE = existingE.concat(errorsArr);
        DriveApp.getFileById(errorsFileId).setContent(JSON.stringify(mergedE));
      } else {
        var ef = DriveApp.createFile((sh.getRange(pickRow,1).getValue() || 'job') + '_errors.json', JSON.stringify(errorsArr));
        sh.getRange(pickRow, 8).setValue(ef.getId());
      }
    }

    // finalize if complete
    var newDone = Number(sh.getRange(pickRow, 5).getValue()) || 0;
    if (newDone >= total) {
      var createdIdNow = sh.getRange(pickRow, 9).getValue() || '';
      var createdAll = [];
      try { createdAll = JSON.parse(DriveApp.getFileById(createdIdNow).getBlob().getDataAsString()) || []; } catch (e) { createdAll = []; }
      var ownerEmail = (sh.getRange(pickRow, 10).getValue() || '').toString();
      try {
        var reportUrl = createReportSpreadsheet(createdAll);
        sh.getRange(pickRow, 7).setValue(reportUrl);
        sh.getRange(pickRow, 4).setValue('done');
        // Notify job owner via email
        try {
          if (ownerEmail) {
            var subject = 'Schoogle: Job ' + (sh.getRange(pickRow,1).getValue() || '') + ' completed';
            var body = 'Your Schoogle job has completed.\n\nCreated events: ' + (createdAll.length || 0) + '\nReport: ' + reportUrl + '\n\nIf you did not expect this, contact the administrator.';
            MailApp.sendEmail(ownerEmail, subject, body);
          }
        } catch (mailErr) {
          console.error('MailApp.sendEmail failed', mailErr);
        }
      } catch (finalErr) {
        sh.getRange(pickRow, 4).setValue('error');
        var ef2 = DriveApp.createFile((sh.getRange(pickRow,1).getValue() || 'job') + '_final_error.txt', finalErr.toString());
        sh.getRange(pickRow, 8).setValue(ef2.getId());
        // Notify owner of failure
        try {
          if (ownerEmail) {
            var subject2 = 'Schoogle: Job ' + (sh.getRange(pickRow,1).getValue() || '') + ' failed';
            var body2 = 'Your Schoogle job encountered an error during finalization.\n\nError: ' + finalErr.toString() + '\nPlease check the job log.';
            MailApp.sendEmail(ownerEmail, subject2, body2);
          }
        } catch (mailErr2) {
          console.error('MailApp.sendEmail failed (error notification)', mailErr2);
        }
      }
    }
  } finally {
    lock.releaseLock();
  }
}

function createReportSpreadsheet(reportData) {
  // reportData is array of created event rows with keys used above
  var ss = SpreadsheetApp.create('Schoogle Timetable');
  var url = ss.getUrl();

  // Created Events sheet
  var sheet = ss.getActiveSheet();
  sheet.setName('Created Events');
  var headers = ['Title', 'Start', 'End', 'Location', 'Google EventID', 'CalID', 'Day', 'Period', 'DayType', 'SchoogleEventID'];
  sheet.getRange(1, 1, 1, headers.length).setValues([headers]);
  if (reportData && reportData.length) {
    var values = reportData.map(function (r) {
      return [r.Title, r.Start, r.End, r.Location, r.GoogleEventID, r.CalID, r.Day, r.Period, r.DayType, r.SchoogleEventID];
    });
    sheet.getRange(2, 1, values.length, headers.length).setValues(values);
    sheet.autoResizeColumns(1, headers.length);
  }

  // Class Sessions sheet
  var cs = ss.insertSheet('Class Sessions');
  // Collect unique titles
  var byTitle = {};
  (reportData || []).forEach(function (r) {
    var t = r.Title || '';
    if (!t) return;
    if (!byTitle[t]) byTitle[t] = [];
    byTitle[t].push(new Date(r.Start));
  });
  var titles = Object.keys(byTitle).sort();
  // Prepare columns: each title is a column with sorted times beneath
  var maxRows = 0;
  titles.forEach(function (t) {
    byTitle[t].sort(function (a, b) { return a.getTime() - b.getTime(); });
    if (byTitle[t].length > maxRows) maxRows = byTitle[t].length;
  });
  if (titles.length) {
    // Headers
    cs.getRange(1, 1, 1, titles.length).setValues([titles]);
    // Columns data as rows
    for (var r = 0; r < maxRows; r++) {
      var rowVals = [];
      for (var c = 0; c < titles.length; c++) {
        rowVals.push(byTitle[titles[c]][r] || '');
      }
      cs.getRange(2 + r, 1, 1, titles.length).setValues([rowVals]);
    }
    cs.autoResizeColumns(1, titles.length);
  }

  return url;
}

function getEventsFromCalendar(calendarId) {
  var out = [];
  var pageToken;
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
