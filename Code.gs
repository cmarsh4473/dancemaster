/**
 * ============================================================================
 * DANCE STUDIO CALENDAR MANAGER v2
 * ============================================================================
 */

var CONFIG = {
  CLASSES_SHEET: "Classes",
  ROLES_SHEET: "Roles",
  STUDIO_CONFIG_SHEET: "Studio Config",
  INSTRUCTORS_SHEET: "Instructors",
  STUDENTS_SHEET: "Students",
  RECURRENCE_YEARS: 10,
  IGNORE_PATTERN: /\bLevel\s*[4-7]\b/i,
  DEFAULT_ROLES: ["Snowflakes", "Rat King", "Party People", "Mother Ginger", "Other"],
  
  STUDIOS: [
    { name: "Charlotte New",      colorId: "7",  location: "Charlotte" },
    { name: "Charlotte Large",    colorId: "2",  location: "Charlotte" },
    { name: "Charlotte Blue",     colorId: "9",  location: "Charlotte" },
    { name: "Charlotte Pink",     colorId: "4",  location: "Charlotte" },
    { name: "Concord Palmetto",   colorId: "5",  location: "Concord" },
    { name: "Concord Magnolia",   colorId: "6",  location: "Concord" },
    { name: "Concord Azalea",     colorId: "11", location: "Concord" },
    { name: "Concord Primrose",   colorId: "10", location: "Concord" }
  ]
};

// 23 columns (A=0 ... W=22)
var COL = {
  SEASON: 0, NAME: 1, TYPE: 2, DAY: 3, START_TIME: 4, LENGTH: 5,
  INSTRUCTOR: 6, ASSISTANTS: 7, LOCATION: 8, MAX_SIZE: 9,
  ENROLLMENT_COUNT: 10, ENROLLMENT_PCT: 11, WAITLIST_COUNT: 12,
  SESSION_CHARGE: 13, ALLOW_DROPINS: 14,
  STUDIO: 15, INSTRUCTOR_EMAIL: 16,
  EVENT_ID: 17, CALENDAR_ID: 18, STATUS: 19,
  SYNC_HASH: 20, ERROR_LOG: 21, STUDENTS: 22
};

var STUDENT_COL = {
  NAME: 0, EMAIL: 1, ROLE: 2, NOTES: 3
};

var INSTRUCTOR_COL = {
  NAME: 0, EMAIL: 1, PHONE: 2, NOTES: 3
};

function onOpen() {
  SpreadsheetApp.getUi()
    .createMenu("Dance Studio")
    .addItem("1. Setup Spreadsheet", "setupSpreadsheet")
    .addItem("2. Setup Calendars", "setupCalendars")
    .addItem("3. Sync to Calendars", "syncCalendars")
    .addSeparator()
    .addItem("4. Clear All Events (Danger)", "clearAllEvents")
    .addToUi();
}

function setupSpreadsheet() {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var classesSheet = ss.getSheetByName(CONFIG.CLASSES_SHEET);
  if (!classesSheet) classesSheet = ss.insertSheet(CONFIG.CLASSES_SHEET);
  
  var headers = [
    "Season", "Name", "Type", "Day", "Start Time", "Length", "Instructor", 
    "Assistants", "Location", "Max Size", "Enrollment Count", "Enrollment %", 
    "Waitlist Count", "Session Charge", "Allow Drop-ins",
    "Studio", "Instructor Email", "Event ID", "Calendar ID", "Status",
    "Sync Hash", "Error Log", "Students"
  ];
  
  classesSheet.getRange(1, 1, 1, headers.length).setValues([headers])
    .setFontWeight("bold").setBackground("#4285f4").setFontColor("white");
  
  var studioNames = CONFIG.STUDIOS.map(function(s) { return s.name; });
  classesSheet.getRange("P2:P").setDataValidation(
    SpreadsheetApp.newDataValidation().requireValueInList(studioNames, true).setAllowInvalid(false).build()
  );
  
  classesSheet.getRange("T2:T").setDataValidation(
    SpreadsheetApp.newDataValidation().requireValueInList(["Active","Deleted","Force Recreate"], true).build()
  ).setNote("Active = normal. Deleted = remove events. Force Recreate = rebuild next sync.");
  
  classesSheet.getRange("R2:W").setBackground("#f3f3f3").setFontColor("#666666")
    .setNote("Script-managed. Do not edit manually.");
  classesSheet.setFrozenRows(1);
  classesSheet.autoResizeColumns(1, headers.length);
  
  // Studio Config
  var configSheet = ss.getSheetByName(CONFIG.STUDIO_CONFIG_SHEET);
  if (!configSheet) configSheet = ss.insertSheet(CONFIG.STUDIO_CONFIG_SHEET);
  var configHeaders = ["Studio Name", "Color ID", "Calendar ID", "Location", "Color Name"];
  configSheet.getRange(1, 1, 1, configHeaders.length).setValues([configHeaders])
    .setFontWeight("bold").setBackground("#34a853").setFontColor("white");
  
  var colorNames = {"1":"Lavender","2":"Sage","3":"Grape","4":"Flamingo","5":"Banana","6":"Tangerine","7":"Peacock","8":"Graphite","9":"Blueberry","10":"Basil","11":"Tomato"};
  var configData = CONFIG.STUDIOS.map(function(s) {
    return [s.name, s.colorId, "", s.location, colorNames[s.colorId] || "Custom"];
  });
  configSheet.getRange(2, 1, configData.length, configData[0].length).setValues(configData);
  configSheet.autoResizeColumns(1, 5);
  
  // Students sheet
  var studentsSheet = ss.getSheetByName(CONFIG.STUDENTS_SHEET);
  if (!studentsSheet) studentsSheet = ss.insertSheet(CONFIG.STUDENTS_SHEET);
  var studentHeaders = ["Name", "Email", "Role", "Notes"];
  studentsSheet.getRange(1, 1, 1, studentHeaders.length).setValues([studentHeaders])
    .setFontWeight("bold").setBackground("#ea4335").setFontColor("white");
  
  var roles = ["Dancer", "Parent", "Instructor", "Admin", "Other"];
  studentsSheet.getRange("C2:C").setDataValidation(
    SpreadsheetApp.newDataValidation().requireValueInList(roles, true).build()
  );
  studentsSheet.setFrozenRows(1);
  studentsSheet.autoResizeColumns(1, 4);
  
    // Roles sheet
  var rolesSheet = ss.getSheetByName(CONFIG.ROLES_SHEET);
  var sheet3 = ss.getSheetByName("Sheet3");
  
  // Auto-fix if Sheet3 exists from a prior run
  if (sheet3 && !rolesSheet) {
    sheet3.setName(CONFIG.ROLES_SHEET);
    rolesSheet = sheet3;
  } else if (sheet3 && rolesSheet) {
    ss.deleteSheet(sheet3); // Remove stray Sheet3
  }
  
  if (!rolesSheet) {
    rolesSheet = ss.insertSheet(CONFIG.ROLES_SHEET);
    rolesSheet.getRange(1, 1).setValue("Role Name")
      .setFontWeight("bold").setBackground("#ea4335").setFontColor("white");
    var roleData = CONFIG.DEFAULT_ROLES.map(function(r) { return [r]; });
    rolesSheet.getRange(2, 1, roleData.length, 1).setValues(roleData);
    rolesSheet.autoResizeColumns(1, 1);
  }
  
  // Instructors sheet
  var instructorsSheet = ss.getSheetByName(CONFIG.INSTRUCTORS_SHEET);
  if (!instructorsSheet) {
    instructorsSheet = ss.insertSheet(CONFIG.INSTRUCTORS_SHEET);
    var instructorHeaders = ["Name", "Email", "Phone", "Notes"];
    instructorsSheet.getRange(1, 1, 1, instructorHeaders.length).setValues([instructorHeaders])
      .setFontWeight("bold").setBackground("#fbbc04").setFontColor("white");
    instructorsSheet.setFrozenRows(1);
    instructorsSheet.autoResizeColumns(1, 4);
  }

  updateStudentValidation();
  SpreadsheetApp.getUi().alert("Setup Complete", "23-column Classes sheet + Students database ready.", SpreadsheetApp.getUi().ButtonSet.OK);
}

function setupCalendars() {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var configSheet = ss.getSheetByName(CONFIG.STUDIO_CONFIG_SHEET);
  var data = configSheet.getRange(2, 1, CONFIG.STUDIOS.length, 4).getValues();
  for (var i = 0; i < data.length; i++) {
    var studioName = data[i][0], existingId = data[i][2], calendar = null;
    if (existingId) { try { calendar = CalendarApp.getCalendarById(existingId); } catch(e) { calendar = null; } }
    if (!calendar) {
      calendar = CalendarApp.createCalendar(studioName, { summary: "Dance Studio - " + studioName, timeZone: Session.getScriptTimeZone() });
      configSheet.getRange(i + 2, 3).setValue(calendar.getId());
    }
  }
  SpreadsheetApp.getUi().alert("Calendars Ready", "All studios linked. Paste CSV and run Sync.", SpreadsheetApp.getUi().ButtonSet.OK);
}

function syncCalendars() {
  var ui = SpreadsheetApp.getUi();
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var classesSheet = ss.getSheetByName(CONFIG.CLASSES_SHEET);
  var configSheet = ss.getSheetByName(CONFIG.STUDIO_CONFIG_SHEET);
  
  var response = ui.prompt("Season Start Date", "Enter season start (MM/DD/YYYY):", ui.ButtonSet.OK_CANCEL);
  if (response.getSelectedButton() != ui.Button.OK) return;
  var seasonStart = new Date(response.getResponseText().trim());
  if (isNaN(seasonStart.getTime())) seasonStart = new Date();
  
  var studioMap = {};
  var configLastRow = configSheet.getLastRow();
  if (configLastRow >= 2) {
    var configData = configSheet.getRange(2, 1, configLastRow - 1, 5).getValues();
    for (var i = 0; i < configData.length; i++) {
      var name = configData[i][0] ? configData[i][0].toString().trim() : "";
      var calId = configData[i][2] ? configData[i][2].toString().trim() : "";
      if (name) studioMap[name] = { colorId: configData[i][1], calendarId: calId };
    }
  }
  
  var lastRow = classesSheet.getLastRow();
  if (lastRow < 2) { ui.alert("No data found."); return; }
  
  var dataRange = classesSheet.getRange(2, 1, lastRow - 1, 23);
  var rows = dataRange.getValues();
  var updates = [];
  var stats = { created: 0, updated: 0, deleted: 0, skipped: 0, errors: 0, unchanged: 0 };
  
  for (var i = 0; i < rows.length; i++) {
    var row = rows[i];
    var rowNum = i + 2;
    var errorLog = "";
    
    try {
      var name = row[COL.NAME];
      if (!name || name.toString().trim() === "") continue;
      
      if (CONFIG.IGNORE_PATTERN.test(name)) {
        stats.skipped++;
        updates.push({ row: rowNum, errorLog: "" });
        continue;
      }
      
      var studio = row[COL.STUDIO] ? row[COL.STUDIO].toString().trim() : "";
      var status = row[COL.STATUS] ? row[COL.STATUS].toString().trim() : "Active";
      var eventIds = row[COL.EVENT_ID] ? row[COL.EVENT_ID].toString().trim() : "";
      var storedCalId = row[COL.CALENDAR_ID] ? row[COL.CALENDAR_ID].toString().trim() : "";
      var storedHash = row[COL.SYNC_HASH] ? row[COL.SYNC_HASH].toString() : "";
      
      if (status === "Deleted") {
        if (eventIds) {
          var delResult = deleteEventSeries(eventIds, storedCalId);
          if (!delResult.success) errorLog += "Delete old: " + delResult.message + "; ";
          stats.deleted++;
        }
        updates.push({ row: rowNum, eventId: "", calId: "", status: "Deleted", syncHash: "", errorLog: errorLog });
        continue;
      }
      
      if (!studio) {
        stats.skipped++;
        updates.push({ row: rowNum, errorLog: "" });
        continue;
      }
      
      var studioInfo = studioMap[studio];
      if (!studioInfo || !studioInfo.calendarId) {
        errorLog += "Studio not configured. ";
        stats.errors++;
        updates.push({ row: rowNum, errorLog: errorLog });
        continue;
      }
      
      var calendar = CalendarApp.getCalendarById(studioInfo.calendarId);
      if (!calendar) {
        errorLog += "Cannot access calendar. ";
        stats.errors++;
        updates.push({ row: rowNum, errorLog: errorLog });
        continue;
      }
      
      var timeSlots = parseClassTimes(row);
      if (!timeSlots || timeSlots.length === 0) {
        errorLog += "Could not parse times. ";
        stats.errors++;
        updates.push({ row: rowNum, errorLog: errorLog });
        continue;
      }
      
      var currentHash = computeEventHash(row, timeSlots);
      var needsRecreate = false;
      var needsGuestUpdate = false;
      
      if (status === "Force Recreate") {
        needsRecreate = true;
      } else if (!eventIds) {
        needsRecreate = true;
      } else if (storedCalId !== studioInfo.calendarId) {
        needsRecreate = true;
      } else if (storedHash !== currentHash) {
        needsRecreate = true;
      } else {
        var ids = eventIds.split(",");
        if (ids.length !== timeSlots.length) {
          needsRecreate = true;
        } else {
          try {
            var testEvent = calendar.getEventSeriesById(ids[0].trim());
            if (!testEvent) needsRecreate = true;
          } catch (e) {
            needsRecreate = true;
          }
        }
        var instructorEmail = row[COL.INSTRUCTOR_EMAIL] ? row[COL.INSTRUCTOR_EMAIL].toString().trim() : "";
        if (instructorEmail && !needsRecreate) needsGuestUpdate = true;
      }
      
      if (needsRecreate) {
        if (eventIds) {
          var delResult = deleteEventSeries(eventIds, storedCalId);
          if (!delResult.success) errorLog += "Delete old: " + delResult.message + "; ";
        }
        var newEventIds = createEventSeriesForRow(row, timeSlots, calendar, studioInfo.colorId, seasonStart);
        updates.push({
          row: rowNum, eventId: newEventIds.join(","), calId: studioInfo.calendarId,
          status: "Active", syncHash: currentHash, errorLog: errorLog
        });
        stats.created++;
        continue;
      }
      
      if (needsGuestUpdate) {
        var instructorEmail = row[COL.INSTRUCTOR_EMAIL].toString().trim();
        var ids = eventIds.split(",");
        for (var j = 0; j < ids.length; j++) {
          try {
            var series = calendar.getEventSeriesById(ids[j].trim());
            if (series) {
              var guests = series.getGuestList();
              for (var g = 0; g < guests.length; g++) series.removeGuest(guests[g].getEmail());
              if (instructorEmail) series.addGuest(instructorEmail);
            }
          } catch (e) {
            errorLog += "Guest update failed: " + e.message + "; ";
          }
        }
        updates.push({ row: rowNum, errorLog: errorLog });
        stats.updated++;
        continue;
      }
      
      updates.push({ row: rowNum, errorLog: errorLog });
      stats.unchanged++;
      
    } catch (rowError) {
      errorLog += "FATAL: " + rowError.message + "; ";
      stats.errors++;
      updates.push({ row: rowNum, errorLog: errorLog });
      Logger.log("Row " + rowNum + " fatal: " + rowError);
    }
  }
  
  for (var u = 0; u < updates.length; u++) {
    var up = updates[u];
    if (up.eventId !== undefined) classesSheet.getRange(up.row, COL.EVENT_ID + 1).setValue(up.eventId);
    if (up.calId !== undefined) classesSheet.getRange(up.row, COL.CALENDAR_ID + 1).setValue(up.calId);
    if (up.status !== undefined) classesSheet.getRange(up.row, COL.STATUS + 1).setValue(up.status);
    if (up.syncHash !== undefined) classesSheet.getRange(up.row, COL.SYNC_HASH + 1).setValue(up.syncHash);
    classesSheet.getRange(up.row, COL.ERROR_LOG + 1).setValue(up.errorLog || "");
  }
  
  var summary = "Created: " + stats.created + "\nUpdated: " + stats.updated + "\nDeleted: " + stats.deleted +
    "\nSkipped: " + stats.skipped + "\nUnchanged: " + stats.unchanged + "\nErrors: " + stats.errors;
  if (stats.errors > 0) summary += "\n\nCheck the 'Error Log' column (V) for details.";
  ui.alert("Sync Complete", summary, ui.ButtonSet.OK);
}

function computeEventHash(row, timeSlots) {
  var parts = [
    row[COL.STUDIO] || "",
    extractClassTitle(row[COL.NAME]),
    row[COL.INSTRUCTOR_EMAIL] || "",
    row[COL.STUDENTS] || ""
  ];
  for (var i = 0; i < timeSlots.length; i++) {
    var t = timeSlots[i];
    parts.push(t.weekday + "|" + t.startHour + ":" + t.startMinute + "|" + t.durationMinutes);
  }
  return Utilities.computeDigest(Utilities.DigestAlgorithm.MD5, parts.join("::"))
    .map(function(b) { return (b < 0 ? b + 256 : b).toString(16).padStart(2, '0'); }).join('');
}

function getInstructors() {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var sheet = ss.getSheetByName(CONFIG.INSTRUCTORS_SHEET);
  if (!sheet) return [];
  var lastRow = sheet.getLastRow();
  if (lastRow < 2) return [];
  
  var data = sheet.getRange(2, 1, lastRow - 1, 4).getValues();
  var instructors = [];
  for (var i = 0; i < data.length; i++) {
    if (!data[i][INSTRUCTOR_COL.NAME]) continue;
    instructors.push({
      rowNum: i + 2,
      name: data[i][INSTRUCTOR_COL.NAME].toString().trim(),
      email: data[i][INSTRUCTOR_COL.EMAIL] ? data[i][INSTRUCTOR_COL.EMAIL].toString().trim() : "",
      phone: data[i][INSTRUCTOR_COL.PHONE] ? data[i][INSTRUCTOR_COL.PHONE].toString().trim() : "",
      notes: data[i][INSTRUCTOR_COL.NOTES] ? data[i][INSTRUCTOR_COL.NOTES].toString() : ""
    });
  }
  return instructors;
}

function addInstructor(data) {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var sheet = ss.getSheetByName(CONFIG.INSTRUCTORS_SHEET);
  var row = sheet.getLastRow() + 1;
  sheet.getRange(row, 1, 1, 4).setValues([[
    data.name,
    data.email,
    data.phone || "",
    data.notes || ""
  ]]);
  return { success: true, row: row };
}

function updateInstructor(data) {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var sheet = ss.getSheetByName(CONFIG.INSTRUCTORS_SHEET);
  sheet.getRange(data.rowNum, 1, 1, 4).setValues([[
    data.name,
    data.email,
    data.phone || "",
    data.notes || ""
  ]]);
  return { success: true };
}

function deleteInstructor(rowNum) {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var sheet = ss.getSheetByName(CONFIG.INSTRUCTORS_SHEET);
  sheet.deleteRow(rowNum);
  return { success: true };
}

function createEventSeriesForRow(row, timeSlots, calendar, colorId, seasonStart) {
  var title = extractClassTitle(row[COL.NAME]);
  var description = buildEventDescription(row);
  var location = row[COL.STUDIO];
  var instructorEmail = row[COL.INSTRUCTOR_EMAIL] ? row[COL.INSTRUCTOR_EMAIL].toString().trim() : "";
  
  // Parse student emails from the Students column (W)
  var studentEmails = [];
  var rawStudents = row[COL.STUDENTS] ? row[COL.STUDENTS].toString().trim() : "";
  if (rawStudents) {
    studentEmails = rawStudents.split(",").map(function(e) { return e.trim(); }).filter(function(e) { return e; });
  }
  
  var eventIds = [];
  var farFuture = new Date(seasonStart);
  farFuture.setFullYear(farFuture.getFullYear() + CONFIG.RECURRENCE_YEARS);
  
  for (var i = 0; i < timeSlots.length; i++) {
    var slot = timeSlots[i];
    var firstDate = getNextWeekdayDate(slot.weekday, slot.startHour, slot.startMinute, seasonStart);
    var endDate = new Date(firstDate.getTime() + slot.durationMinutes * 60000);
    var recurrence = CalendarApp.newRecurrence().addWeeklyRule().until(farFuture);
    var series = calendar.createEventSeries(title, firstDate, endDate, recurrence, {
      description: description, location: location
    });
    series.setColor(colorId.toString());
    
    // Small delay so Calendar API propagates the series before adding guests
    Utilities.sleep(500);
    
    // Add instructor
    if (instructorEmail) {
      try {
        series.addGuest(instructorEmail);
      } catch(e) {
        Logger.log("Failed to add instructor guest " + instructorEmail + ": " + e.message);
      }
    }
    
    // Add students
    for (var s = 0; s < studentEmails.length; s++) {
      try {
        series.addGuest(studentEmails[s]);
        Logger.log("Added student guest: " + studentEmails[s]);
      } catch(e) {
        Logger.log("Failed to add student guest " + studentEmails[s] + ": " + e.message);
      }
    }
    
    eventIds.push(series.getId());
  }
  return eventIds;
}

function deleteEventSeries(eventIdsCsv, calendarId) {
  if (!eventIdsCsv || !calendarId) return { success: true, message: "Nothing to delete" };
  var calendar = null;
  try { calendar = CalendarApp.getCalendarById(calendarId); } catch(e) { return { success: false, message: "Calendar not found: " + calendarId }; }
  if (!calendar) return { success: false, message: "Calendar not accessible: " + calendarId };
  
  var ids = eventIdsCsv.split(",");
  var results = [], anySuccess = false;
  for (var i = 0; i < ids.length; i++) {
    var id = ids[i].trim();
    if (!id) continue;
    try {
      var series = calendar.getEventSeriesById(id);
      if (series) { series.deleteEventSeries(); results.push("OK:" + id.slice(-6)); anySuccess = true; }
      else results.push("Missing:" + id.slice(-6));
    } catch (e) {
      results.push("Err:" + id.slice(-6) + "(" + e.message + ")");
    }
  }
  return { success: anySuccess || results.length === 0, message: results.join(", ") };
}

function parseClassTimes(row) {
  var name = row[COL.NAME];
  var csvDay = row[COL.DAY];
  var csvStartTime = row[COL.START_TIME];
  var csvLength = row[COL.LENGTH];
  var fromName = parseTimesFromName(name, csvStartTime);
  if (fromName && fromName.length > 0) return fromName;
  return parseTimesFromCSV(csvDay, csvStartTime, csvLength);
}

function parseTimesFromName(name, csvStartTime) {
  var timePart = "";
  var locMatch = name.toString().match(/-\s*(CLT|CON):\s*(.+)/);
  if (locMatch) {
    timePart = locMatch[2];
  } else {
    var genericMatch = name.toString().match(/:\s*(.+)/);
    if (genericMatch) timePart = genericMatch[1];
  }
  if (!timePart) return null;
  
  var dayMap = { "Mon":"Monday", "Tue":"Tuesday", "Wed":"Wednesday", "Thur":"Thursday", "Fri":"Friday", "Sat":"Saturday" };
  
  // FIX: Handle both Date objects (from Sheets) and raw strings
  var isPM = false;
  var isAM = false;
  
  if (csvStartTime instanceof Date) {
    // Sheets returns times as Date objects (hour 0-23)
    var hour = csvStartTime.getHours();
    isPM = hour >= 12;
    isAM = hour < 12;
  } else if (csvStartTime) {
    var startTimeStr = csvStartTime.toString().toUpperCase();
    isPM = startTimeStr.indexOf("PM") !== -1;
    isAM = startTimeStr.indexOf("AM") !== -1;
  }
  
  var segments = [];
  var regex = /(Mon|Tue|Wed|Thur|Fri|Sat)\s+(\d{1,2}):(\d{2})\s*-\s*(\d{1,2}):(\d{2})/g;
  var match;
  while ((match = regex.exec(timePart)) !== null) {
    var startHour = parseInt(match[2]);
    var startMin = parseInt(match[3]);
    var endHour = parseInt(match[4]);
    var endMin = parseInt(match[5]);
    
    if (isPM) {
      if (startHour !== 12) startHour += 12;
      if (endHour !== 12) endHour += 12;
    } else if (isAM) {
      if (startHour === 12) startHour = 0;
      if (endHour === 12) endHour = 0;
    }
    
    var startTotal = startHour * 60 + startMin;
    var endTotal = endHour * 60 + endMin;
    var duration = endTotal - startTotal;
    if (duration < 0) duration += 24 * 60;
    
    segments.push({
      weekday: dayMap[match[1]],
      startHour: startHour,
      startMinute: startMin,
      durationMinutes: duration
    });
  }
  return segments.length > 0 ? segments : null;
}

function parseTimesFromCSV(day, startTime, length) {
  var dayMap = {
    "Monday":["Monday"], "Tuesday":["Tuesday"], "Wednesday":["Wednesday"],
    "Thursday":["Thursday"], "Friday":["Friday"], "Saturday":["Saturday"],
    "MW":["Monday","Wednesday"], "TuTh":["Tuesday","Thursday"], "TuWTh":["Tuesday","Wednesday","Thursday"]
  };
  var weekdays = dayMap[day];
  if (!weekdays) return null;
  var time = parseTimeString(startTime);
  if (!time) return null;
  var duration = parseLengthString(length);
  if (!duration) return null;
  var segments = [];
  for (var i = 0; i < weekdays.length; i++) {
    segments.push({ weekday: weekdays[i], startHour: time.hour, startMinute: time.minute, durationMinutes: duration });
  }
  return segments;
}

function parseTimeString(timeStr) {
  // Handle Date objects from Google Sheets (times stored as dates)
  if (timeStr instanceof Date) {
    return {
      hour: timeStr.getHours(),
      minute: timeStr.getMinutes()
    };
  }
  
  var match = timeStr.toString().match(/(\d{1,2}):(\d{2})\s*(AM|PM)/i);
  if (!match) return null;
  
  var hour = parseInt(match[1]);
  var minute = parseInt(match[2]);
  var ampm = match[3].toUpperCase();
  
  if (ampm === "PM" && hour !== 12) hour += 12;
  if (ampm === "AM" && hour === 12) hour = 0;
  
  return { hour: hour, minute: minute };
}

function parseLengthString(lengthStr) {
  // Handle Date objects from Google Sheets (e.g. "0:45" stored as 12:45 AM)
  if (lengthStr instanceof Date) {
    return lengthStr.getHours() * 60 + lengthStr.getMinutes();
  }
  
  var parts = lengthStr.toString().split(":");
  var hours = parseInt(parts[0]) || 0;
  var minutes = parseInt(parts[1]) || 0;
  return hours * 60 + minutes;
}


function getNextWeekdayDate(weekdayName, hour, minute, fromDate) {
  var dayMap = { "Sunday":0, "Monday":1, "Tuesday":2, "Wednesday":3, "Thursday":4, "Friday":5, "Saturday":6 };
  var targetDay = dayMap[weekdayName], currentDay = fromDate.getDay(), daysDiff = targetDay - currentDay;
  if (daysDiff < 0) daysDiff += 7;
  if (daysDiff === 0) {
    if (fromDate.getHours() * 60 + fromDate.getMinutes() >= hour * 60 + minute) daysDiff = 7;
  }
  var result = new Date(fromDate);
  result.setDate(result.getDate() + daysDiff);
  result.setHours(hour, minute, 0, 0);
  return result;
}

function extractClassTitle(name) {
  var match = name.toString().match(/^(.+?)\s*-\s*(CLT|CON):/);
  if (match) return match[1].trim();
  return name.toString().trim();
}

function buildEventDescription(row) {
  var parts = [];
  parts.push("Full Class Name: " + row[COL.NAME]);
  parts.push("Location: " + row[COL.LOCATION]);
  parts.push("Max Size: " + (row[COL.MAX_SIZE] || "N/A"));
  parts.push("Enrollment: " + (row[COL.ENROLLMENT_COUNT] || "0") + " / " + (row[COL.MAX_SIZE] || "N/A"));
  if (row[COL.WAITLIST_COUNT]) parts.push("Waitlist: " + row[COL.WAITLIST_COUNT]);
  parts.push("Session Charge: $" + row[COL.SESSION_CHARGE]);
  if (row[COL.INSTRUCTOR] && row[COL.INSTRUCTOR] !== "Staff") parts.push("Instructor: " + row[COL.INSTRUCTOR]);
  if (row[COL.ASSISTANTS]) parts.push("Assistants: " + row[COL.ASSISTANTS]);
  return parts.join("\n");
}

function clearAllEvents() {
  var ui = SpreadsheetApp.getUi();
  if (ui.alert("Delete ALL events?", "This removes every calendar event. Cannot be undone.", ui.ButtonSet.YES_NO) !== ui.Button.YES) return;
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var classesSheet = ss.getSheetByName(CONFIG.CLASSES_SHEET);
  var lastRow = classesSheet.getLastRow();
  if (lastRow < 2) return;
  var data = classesSheet.getRange(2, 1, lastRow - 1, 22).getValues();
  var cleared = 0;
  for (var i = 0; i < data.length; i++) {
    var eventIds = data[i][COL.EVENT_ID] ? data[i][COL.EVENT_ID].toString().trim() : "";
    var calId = data[i][COL.CALENDAR_ID] ? data[i][COL.CALENDAR_ID].toString().trim() : "";
    if (eventIds && calId) {
      deleteEventSeries(eventIds, calId);
      classesSheet.getRange(i + 2, COL.EVENT_ID + 1).setValue("");
      classesSheet.getRange(i + 2, COL.CALENDAR_ID + 1).setValue("");
      classesSheet.getRange(i + 2, COL.SYNC_HASH + 1).setValue("");
      classesSheet.getRange(i + 2, COL.ERROR_LOG + 1).setValue("Cleared");
      cleared++;
    }
  }
  ui.alert("Cleared " + cleared + " event series.");
}

// =============================================================================
// DASHBOARD SERVER FUNCTIONS
// =============================================================================

// function showDashboard() {
//   var html = HtmlService.createHtmlOutputFromFile('Dashboard')
//     .setTitle('Dance Studio Manager')
//     .setWidth(450);
//   SpreadsheetApp.getUi().showSidebar(html);
// }

// =============================================================================
// STUDENT DATABASE FUNCTIONS
// =============================================================================

function getStudents() {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var sheet = ss.getSheetByName(CONFIG.STUDENTS_SHEET);
  if (!sheet) return [];
  var lastRow = sheet.getLastRow();
  if (lastRow < 2) return [];
  
  var data = sheet.getRange(2, 1, lastRow - 1, 4).getValues();
  var students = [];
  for (var i = 0; i < data.length; i++) {
    if (!data[i][STUDENT_COL.NAME]) continue;
    students.push({
      rowNum: i + 2,
      name: data[i][STUDENT_COL.NAME].toString().trim(),
      email: data[i][STUDENT_COL.EMAIL] ? data[i][STUDENT_COL.EMAIL].toString().trim() : "",
      role: data[i][STUDENT_COL.ROLE] ? data[i][STUDENT_COL.ROLE].toString().trim() : "",
      notes: data[i][STUDENT_COL.NOTES] ? data[i][STUDENT_COL.NOTES].toString() : ""
    });
  }
  return students;
}

function addStudent(studentData) {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var sheet = ss.getSheetByName(CONFIG.STUDENTS_SHEET);
  var row = sheet.getLastRow() + 1;
  sheet.getRange(row, 1, 1, 4).setValues([[
    studentData.name,
    studentData.email,
    studentData.role,
    studentData.notes || ""
  ]]);
  return { success: true, row: row };
}

function updateStudent(studentData) {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var sheet = ss.getSheetByName(CONFIG.STUDENTS_SHEET);
  sheet.getRange(studentData.rowNum, 1, 1, 4).setValues([[
    studentData.name,
    studentData.email,
    studentData.role,
    studentData.notes || ""
  ]]);
  return { success: true };
}

function deleteStudent(rowNum) {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var sheet = ss.getSheetByName(CONFIG.STUDENTS_SHEET);
  sheet.deleteRow(rowNum);
  return { success: true };
}

function getDashboardData() {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var classesSheet = ss.getSheetByName(CONFIG.CLASSES_SHEET);
  var configSheet = ss.getSheetByName(CONFIG.STUDIO_CONFIG_SHEET);
  
  // Load studios
  var studioData = configSheet.getRange(2, 1, CONFIG.STUDIOS.length, 2).getValues();
  var studios = [];
  var studioColors = {};
  for (var i = 0; i < studioData.length; i++) {
    studios.push({ name: studioData[i][0] });
    studioColors[studioData[i][0]] = getColorHex(studioData[i][1]);
  }
  
  // Load classes
  var lastRow = classesSheet.getLastRow();
  var classes = [];
  if (lastRow >= 2) {
    var data = classesSheet.getRange(2, 1, lastRow - 1, 22).getValues();
    for (var i = 0; i < data.length; i++) {
      var row = data[i];
      if (!row[COL.NAME] || row[COL.NAME].toString().trim() === "") continue;
      classes.push({
        rowNum: i + 2,
        name: row[COL.NAME],
        day: row[COL.DAY],
        startTime: row[COL.START_TIME] ? row[COL.START_TIME].toString() : "",
        studio: row[COL.STUDIO] ? row[COL.STUDIO].toString() : "",
        instructorEmail: row[COL.INSTRUCTOR_EMAIL] ? row[COL.INSTRUCTOR_EMAIL].toString() : "",
        status: row[COL.STATUS] ? row[COL.STATUS].toString() : "Active"
      });
    }
  }
  
  // Get saved season start
  var props = PropertiesService.getDocumentProperties();
  var seasonStart = props.getProperty('seasonStart') || "";
  
  return { classes: classes, studios: studios, studioColors: studioColors, seasonStart: seasonStart };
}

function saveClassChanges(updates) {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var classesSheet = ss.getSheetByName(CONFIG.CLASSES_SHEET);
  
  for (var i = 0; i < updates.length; i++) {
    var up = updates[i];
    classesSheet.getRange(up.row, COL.STUDIO + 1).setValue(up.studio);
    classesSheet.getRange(up.row, COL.INSTRUCTOR_EMAIL + 1).setValue(up.instructorEmail);
    classesSheet.getRange(up.row, COL.STATUS + 1).setValue(up.status);
    
    // If status changed to Deleted, clear event IDs so sync removes them
    if (up.status === "Deleted") {
      // Keep the old IDs so deleteEventSeries can find them during sync
      // The sync function handles deletion
    }
    
    // If studio changed, mark for recreation by clearing hash
    classesSheet.getRange(up.row, COL.SYNC_HASH + 1).setValue("");
  }
  
  return { success: true, count: updates.length };
}

function addNewClass(classData) {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var classesSheet = ss.getSheetByName(CONFIG.CLASSES_SHEET);
  
  var lastRow = classesSheet.getLastRow();
  var newRow = lastRow + 1;
  
  // Build the row array (23 columns A-W)
  var row = new Array(23).fill("");
  row[COL.SEASON] = "Manual Add";
  row[COL.NAME] = classData.name;
  row[COL.TYPE] = "Downpayment";
  row[COL.DAY] = classData.day;
  row[COL.START_TIME] = classData.startTime;
  row[COL.LENGTH] = classData.length;
  row[COL.INSTRUCTOR] = "Staff";
  row[COL.LOCATION] = classData.location;
  row[COL.MAX_SIZE] = classData.maxSize || "";
  row[COL.SESSION_CHARGE] = "149.00";
  row[COL.ALLOW_DROPINS] = "No";
  row[COL.STUDIO] = classData.studio;
  row[COL.INSTRUCTOR_EMAIL] = classData.instructorEmail || "";
  row[COL.STATUS] = "Active";
  row[COL.STUDENTS] = ""; // Explicitly blank for new classes
  
  classesSheet.getRange(newRow, 1, 1, 23).setValues([row]);
  
  return { success: true, row: newRow };
}

function getRoles() {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var sheet = ss.getSheetByName(CONFIG.ROLES_SHEET);
  
  if (!sheet) {
    Logger.log("WARNING: Roles sheet '" + CONFIG.ROLES_SHEET + "' not found. Using defaults.");
    return CONFIG.DEFAULT_ROLES;
  }
  
  var lastRow = sheet.getLastRow();
  if (lastRow < 2) {
    Logger.log("WARNING: Roles sheet exists but is empty. Using defaults.");
    return CONFIG.DEFAULT_ROLES;
  }
  
  var data = sheet.getRange(2, 1, lastRow - 1, 1).getValues();
  var roles = data.filter(function(r) { return r[0]; }).map(function(r) { return r[0].toString().trim(); });
  
  Logger.log("Loaded roles: " + roles.join(", "));
  return roles.length > 0 ? roles : CONFIG.DEFAULT_ROLES;
}

function updateStudentValidation() {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var sheet = ss.getSheetByName(CONFIG.STUDENTS_SHEET);
  var roles = getRoles();
  if (roles.length > 0 && sheet) {
    sheet.getRange("C2:C").setDataValidation(
      SpreadsheetApp.newDataValidation().requireValueInList(roles, true).build()
    );
  }
}

function addRole(roleName) {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var sheet = ss.getSheetByName(CONFIG.ROLES_SHEET);
  var lastRow = sheet.getLastRow();
  sheet.getRange(lastRow + 1, 1).setValue(roleName);
  updateStudentValidation();
  return { success: true, row: lastRow + 1 };
}

function deleteRole(rowNum) {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var sheet = ss.getSheetByName(CONFIG.ROLES_SHEET);
  sheet.deleteRow(rowNum);
  updateStudentValidation();
  return { success: true };
}

function updateRole(rowNum, roleName) {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var sheet = ss.getSheetByName(CONFIG.ROLES_SHEET);
  sheet.getRange(rowNum, 1).setValue(roleName);
  updateStudentValidation();
  return { success: true };
}

function quickSync(seasonStartStr) {
  PropertiesService.getDocumentProperties().setProperty('seasonStart', seasonStartStr);
  
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var classesSheet = ss.getSheetByName(CONFIG.CLASSES_SHEET);
  var configSheet = ss.getSheetByName(CONFIG.STUDIO_CONFIG_SHEET);
  
  var seasonStart = new Date(seasonStartStr);
  if (isNaN(seasonStart.getTime())) {
    throw new Error("Invalid season start date: " + seasonStartStr);
  }
  
  var studioMap = {};
  var configLastRow = configSheet.getLastRow();
  if (configLastRow >= 2) {
    var configData = configSheet.getRange(2, 1, configLastRow - 1, 5).getValues();
    for (var i = 0; i < configData.length; i++) {
      var name = configData[i][0] ? configData[i][0].toString().trim() : "";
      var calId = configData[i][2] ? configData[i][2].toString().trim() : "";
      if (name) studioMap[name] = { colorId: configData[i][1], calendarId: calId };
    }
  }
  
  var lastRow = classesSheet.getLastRow();
  if (lastRow < 2) throw new Error("No data found in Classes sheet.");
  
  var dataRange = classesSheet.getRange(2, 1, lastRow - 1, 23);
  var rows = dataRange.getValues();
  var updates = [];
  var stats = { created: 0, updated: 0, deleted: 0, skipped: 0, errors: 0, unchanged: 0 };
  
  for (var i = 0; i < rows.length; i++) {
    var row = rows[i];
    var rowNum = i + 2;
    var errorLog = "";
    
    try {
      var name = row[COL.NAME];
      if (!name || name.toString().trim() === "") continue;
      
      if (CONFIG.IGNORE_PATTERN.test(name)) {
        stats.skipped++;
        updates.push({ row: rowNum, errorLog: "" });
        continue;
      }
      
      var studio = row[COL.STUDIO] ? row[COL.STUDIO].toString().trim() : "";
      var status = row[COL.STATUS] ? row[COL.STATUS].toString().trim() : "Active";
      var eventIds = row[COL.EVENT_ID] ? row[COL.EVENT_ID].toString().trim() : "";
      var storedCalId = row[COL.CALENDAR_ID] ? row[COL.CALENDAR_ID].toString().trim() : "";
      var storedHash = row[COL.SYNC_HASH] ? row[COL.SYNC_HASH].toString() : "";
      
      if (status === "Deleted") {
        if (eventIds) {
          var delResult = deleteEventSeries(eventIds, storedCalId);
          if (!delResult.success) errorLog += delResult.message + "; ";
          stats.deleted++;
        }
        updates.push({ row: rowNum, eventId: "", calId: "", status: "Deleted", syncHash: "", errorLog: errorLog });
        continue;
      }
      
      if (!studio) {
        stats.skipped++;
        updates.push({ row: rowNum, errorLog: "" });
        continue;
      }
      
      var studioInfo = studioMap[studio];
      if (!studioInfo || !studioInfo.calendarId) {
        errorLog += "Studio not configured. ";
        stats.errors++;
        updates.push({ row: rowNum, errorLog: errorLog });
        continue;
      }
      
      var calendar = CalendarApp.getCalendarById(studioInfo.calendarId);
      if (!calendar) {
        errorLog += "Cannot access calendar. ";
        stats.errors++;
        updates.push({ row: rowNum, errorLog: errorLog });
        continue;
      }
      
      var timeSlots = parseClassTimes(row);
      if (!timeSlots || timeSlots.length === 0) {
        var rawDay = row[COL.DAY] || "(blank)";
        var rawStart = row[COL.START_TIME] ? row[COL.START_TIME].toString() : "(blank)";
        var rawLength = row[COL.LENGTH] ? row[COL.LENGTH].toString() : "(blank)";
        errorLog += "Time parse failed. Name has no times. CSV columns → Day: '" + rawDay + "', Start: '" + rawStart + "', Length: '" + rawLength + "'. ";
        stats.errors++;
        updates.push({ row: rowNum, errorLog: errorLog });
        continue;
      }
      
      var currentHash = computeEventHash(row, timeSlots);
      var needsRecreate = false;
      var needsGuestUpdate = false;
      
      if (status === "Force Recreate") {
        needsRecreate = true;
      } else if (!eventIds) {
        needsRecreate = true;
      } else if (storedCalId !== studioInfo.calendarId) {
        needsRecreate = true;
      } else if (storedHash !== currentHash) {
        needsRecreate = true;
      } else {
        var ids = eventIds.split(",");
        if (ids.length !== timeSlots.length) {
          needsRecreate = true;
        } else {
          try {
            var testEvent = calendar.getEventSeriesById(ids[0].trim());
            if (!testEvent) needsRecreate = true;
          } catch (e) {
            needsRecreate = true;
          }
        }
        var instructorEmail = row[COL.INSTRUCTOR_EMAIL] ? row[COL.INSTRUCTOR_EMAIL].toString().trim() : "";
        if (instructorEmail && !needsRecreate) needsGuestUpdate = true;
      }
      
      if (needsRecreate) {
        if (eventIds) {
          var delResult = deleteEventSeries(eventIds, storedCalId);
          if (!delResult.success) errorLog += "Old delete: " + delResult.message + "; ";
          if (storedCalId && storedCalId !== studioInfo.calendarId) {
            var delResult2 = deleteEventSeries(eventIds, studioInfo.calendarId);
            if (!delResult2.success) errorLog += "New delete: " + delResult2.message + "; ";
          }
        }
        var newEventIds = createEventSeriesForRow(row, timeSlots, calendar, studioInfo.colorId, seasonStart);
        updates.push({
          row: rowNum, eventId: newEventIds.join(","), calId: studioInfo.calendarId,
          status: "Active", syncHash: currentHash, errorLog: errorLog
        });
        stats.created++;
        continue;
      }
      
      if (needsGuestUpdate) {
        var instructorEmail = row[COL.INSTRUCTOR_EMAIL].toString().trim();
        var ids = eventIds.split(",");
        for (var j = 0; j < ids.length; j++) {
          try {
            var series = calendar.getEventSeriesById(ids[j].trim());
            if (series) {
              var guests = series.getGuestList();
              for (var g = 0; g < guests.length; g++) series.removeGuest(guests[g].getEmail());
              if (instructorEmail) series.addGuest(instructorEmail);
            }
          } catch (e) {
            errorLog += "Guest update failed: " + e.message + "; ";
          }
        }
        updates.push({ row: rowNum, errorLog: errorLog });
        stats.updated++;
        continue;
      }
      
      updates.push({ row: rowNum, errorLog: errorLog });
      stats.unchanged++;
      
    } catch (rowError) {
      errorLog += "FATAL: " + rowError.message + "; ";
      stats.errors++;
      updates.push({ row: rowNum, errorLog: errorLog });
      Logger.log("Row " + rowNum + " fatal: " + rowError);
    }
  }
  
  for (var u = 0; u < updates.length; u++) {
    var up = updates[u];
    if (up.eventId !== undefined) classesSheet.getRange(up.row, COL.EVENT_ID + 1).setValue(up.eventId);
    if (up.calId !== undefined) classesSheet.getRange(up.row, COL.CALENDAR_ID + 1).setValue(up.calId);
    if (up.status !== undefined) classesSheet.getRange(up.row, COL.STATUS + 1).setValue(up.status);
    if (up.syncHash !== undefined) classesSheet.getRange(up.row, COL.SYNC_HASH + 1).setValue(up.syncHash);
    classesSheet.getRange(up.row, COL.ERROR_LOG + 1).setValue(up.errorLog || "");
  }
  
  return {
    success: true,
    message: "Sync complete! Created: " + stats.created + ", Updated: " + stats.updated + ", Deleted: " + stats.deleted + ", Errors: " + stats.errors
  };
}

function getColorHex(colorId) {
  var map = {
    "1": "#7986cb", "2": "#33b679", "3": "#8e24aa", "4": "#e67c73",
    "5": "#f6c026", "6": "#f5511d", "7": "#039be5", "8": "#616161",
    "9": "#3f51b5", "10": "#0b8043", "11": "#d50000"
  };
  return map[colorId] || "#999999";
}

// Add dashboard to menu
function onOpen() {
  SpreadsheetApp.getUi()
    .createMenu("Dance Studio")
    .addSeparator()
    .addItem("1. Setup Spreadsheet", "setupSpreadsheet")
    .addItem("2. Setup Calendars", "setupCalendars")
    .addItem("3. Sync to Calendars", "syncCalendars")
    .addSeparator()
    .addItem("4. Clear All Events (Danger)", "clearAllEvents")
    .addToUi();
}

// =============================================================================
// WEB APP ENTRY POINTS
// =============================================================================

function doGet(e) {
  return HtmlService.createHtmlOutputFromFile('Index')
    .setTitle('Dance Studio Manager')
    .setXFrameOptionsMode(HtmlService.XFrameOptionsMode.ALLOWALL);
}

function getWebAppData() {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var classesSheet = ss.getSheetByName(CONFIG.CLASSES_SHEET);
  var configSheet = ss.getSheetByName(CONFIG.STUDIO_CONFIG_SHEET);
  
  // Studios
  var studioData = configSheet.getRange(2, 1, CONFIG.STUDIOS.length, 5).getValues();
  var studios = [];
  var studioColors = {};
  for (var i = 0; i < studioData.length; i++) {
    studios.push({ 
      name: studioData[i][0], 
      colorId: studioData[i][1], 
      location: studioData[i][3],
      colorName: studioData[i][4] 
    });
    studioColors[studioData[i][0]] = getColorHex(studioData[i][1]);
  }
  
  // Classes
  var lastRow = classesSheet.getLastRow();
  var classes = [];
  if (lastRow >= 2) {
    var data = classesSheet.getRange(2, 1, lastRow - 1, 23).getValues();
    for (var i = 0; i < data.length; i++) {
      var row = data[i];
      if (!row[COL.NAME] || row[COL.NAME].toString().trim() === "") continue;
      
      var rawStudents = row[COL.STUDENTS] ? row[COL.STUDENTS].toString().trim() : "";
      var studentList = rawStudents ? rawStudents.split(",").map(function(e){ return e.trim(); }).filter(function(e){ return e; }) : [];
      
      classes.push({
        rowNum: i + 2,
        season: row[COL.SEASON],
        name: row[COL.NAME],
        title: extractClassTitle(row[COL.NAME]),
        day: row[COL.DAY],
        startTime: formatTimeValue(row[COL.START_TIME]),
        length: formatLengthValue(row[COL.LENGTH]),
        instructor: row[COL.INSTRUCTOR],
        location: row[COL.LOCATION],
        maxSize: row[COL.MAX_SIZE],
        enrollment: row[COL.ENROLLMENT_COUNT],
        studio: row[COL.STUDIO] ? row[COL.STUDIO].toString() : "",
        instructorEmail: row[COL.INSTRUCTOR_EMAIL] ? row[COL.INSTRUCTOR_EMAIL].toString() : "",
        students: studentList,
        status: row[COL.STATUS] ? row[COL.STATUS].toString() : "Active",
        eventId: row[COL.EVENT_ID] ? row[COL.EVENT_ID].toString() : "",
        errorLog: row[COL.ERROR_LOG] ? row[COL.ERROR_LOG].toString() : ""
      });
    }
  }
  
  var props = PropertiesService.getDocumentProperties();
  return {
    classes: classes,
    studios: studios,
    studioColors: studioColors,
    seasonStart: props.getProperty('seasonStart') || "",
    students: getStudents(),
    instructors: getInstructors(),
    roles: getRoles()
  };
}

function formatTimeValue(val) {
  if (!val) return "";
  if (val instanceof Date) {
    var h = val.getHours();
    var m = val.getMinutes();
    var ampm = h >= 12 ? 'PM' : 'AM';
    h = h % 12;
    if (h === 0) h = 12;
    return h + ":" + (m < 10 ? "0" + m : m) + " " + ampm;
  }
  return val.toString();
}

function formatLengthValue(val) {
  if (!val) return "";
  if (val instanceof Date) {
    var h = val.getHours();
    var m = val.getMinutes();
    return h + ":" + (m < 10 ? "0" + m : m);
  }
  return val.toString();
}

function webSaveChanges(updates) {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var classesSheet = ss.getSheetByName(CONFIG.CLASSES_SHEET);
  for (var i = 0; i < updates.length; i++) {
    var up = updates[i];
    if (up.day !== undefined) classesSheet.getRange(up.row, COL.DAY + 1).setValue(up.day);
    if (up.startTime !== undefined) classesSheet.getRange(up.row, COL.START_TIME + 1).setValue(up.startTime);
    if (up.length !== undefined) classesSheet.getRange(up.row, COL.LENGTH + 1).setValue(up.length);
    classesSheet.getRange(up.row, COL.STUDIO + 1).setValue(up.studio);
    classesSheet.getRange(up.row, COL.INSTRUCTOR_EMAIL + 1).setValue(up.instructorEmail);
    classesSheet.getRange(up.row, COL.STATUS + 1).setValue(up.status);
    if (up.students !== undefined) {
      var studentStr = Array.isArray(up.students) ? up.students.join(", ") : up.students;
      classesSheet.getRange(up.row, COL.STUDENTS + 1).setValue(studentStr);
    }
    classesSheet.getRange(up.row, COL.SYNC_HASH + 1).setValue("");
  }
  return { success: true, count: updates.length };
}

function webAddClass(classData) {
  return addNewClass(classData);
}

function webDeleteClass(rowNum) {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var classesSheet = ss.getSheetByName(CONFIG.CLASSES_SHEET);
  classesSheet.getRange(rowNum, COL.STATUS + 1).setValue("Deleted");
  classesSheet.getRange(rowNum, COL.SYNC_HASH + 1).setValue("");
  return { success: true, message: "Class removed. Run Sync to delete from calendar." };
}

function webSync(seasonStartStr) {
  return quickSync(seasonStartStr);
}