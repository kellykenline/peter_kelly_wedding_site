const SPREADSHEET_ID = "1hL4kQoY7HDOlotGPbBkvnGJj_lC54gg4AmrkjsHnC2Y";
const SHEET_NAME = "RSVP Responses";

const HEADERS = [
  "Submitted At",
  "Household ID",
  "Lookup First Name",
  "Lookup Last Name",
  "Guest Count",
  "Guest Number",
  "Guest First Name",
  "Guest Last Name",
  "Guest Attendance",
  "Adult or Child",
  "Meal Selection",
  "Attending Rehearsal Dinner",
  "Attending Morning-After Brunch",
  "Dietary Restrictions / Other Comments"
];

function doGet(e) {
  const action = e.parameter.action;
  const callback = e.parameter.callback || "callback";

  if (action === "lookup") {
    const matches = findRsvpsByGuest_(e.parameter.firstName || "", e.parameter.lastName || "");
    const output = JSON.stringify({
      result: "success",
      found: matches.length > 0,
      matches: matches,
      data: matches.length === 1 ? matches[0] : null
    });

    return ContentService
      .createTextOutput(callback + "(" + output + ");")
      .setMimeType(ContentService.MimeType.JAVASCRIPT);
  }

  return ContentService.createTextOutput("RSVP endpoint is active.");
}

function doPost(e) {
  try {
    const data = JSON.parse(e.parameter.payload || "{}");
    const guests = Array.isArray(data.guests) ? data.guests : [];

    if (guests.length === 0) {
      throw new Error("No guest records were received.");
    }

    guests.forEach(function(guest, index) {
      if (!String(guest.firstName || "").trim() || !String(guest.lastName || "").trim()) {
        throw new Error("Guest " + (index + 1) + " is missing a first or last name.");
      }
    });

    const ss = SpreadsheetApp.openById(SPREADSHEET_ID);
    const sheet = getOrCreateSheet_(ss);
    const householdId = data.householdId || Utilities.getUuid();

    removeExistingRowsByHousehold_(sheet, householdId);

    const rows = guests.map(function(guest, index) {
      return [
        data.submittedAt || new Date().toISOString(),
        householdId,
        data.lookupFirstName || "",
        data.lookupLastName || "",
        guests.length,
        guest.guestNumber || (index + 1),
        guest.firstName || "",
        guest.lastName || "",
        guest.status || "Joyfully Accept",
        guest.type || "",
        guest.meal || "",
        guest.rehearsalDinner || "No",
        guest.morningAfterBrunch || "No",
        guest.comments || ""
      ];
    });

    sheet.getRange(sheet.getLastRow() + 1, 1, rows.length, HEADERS.length).setValues(rows);
    SpreadsheetApp.flush();

    return ContentService
      .createTextOutput(JSON.stringify({
        result: "success",
        householdId: householdId,
        guestCount: rows.length
      }))
      .setMimeType(ContentService.MimeType.JSON);
  } catch (error) {
    return ContentService
      .createTextOutput(JSON.stringify({
        result: "error",
        message: error.message
      }))
      .setMimeType(ContentService.MimeType.JSON);
  }
}

function getOrCreateSheet_(ss) {
  let sheet = ss.getSheetByName(SHEET_NAME);
  if (!sheet) sheet = ss.insertSheet(SHEET_NAME);

  if (sheet.getLastRow() === 0) {
    sheet.appendRow(HEADERS);
  }

  return sheet;
}

function normalize_(value) {
  return String(value || "").trim().toLowerCase();
}

function findRsvpsByGuest_(firstName, lastName) {
  const ss = SpreadsheetApp.openById(SPREADSHEET_ID);
  const sheet = getOrCreateSheet_(ss);
  const values = sheet.getDataRange().getValues();

  if (values.length <= 1) return [];

  const targetFirst = normalize_(firstName);
  const targetLast = normalize_(lastName);
  const matchingHouseholds = [];

  for (let i = 1; i < values.length; i++) {
    if (
      normalize_(values[i][6]) === targetFirst &&
      normalize_(values[i][7]) === targetLast
    ) {
      const householdId = String(values[i][1] || "");
      if (householdId && matchingHouseholds.indexOf(householdId) === -1) {
        matchingHouseholds.push(householdId);
      }
    }
  }

  return matchingHouseholds.map(function(householdId) {
    const rows = [];

    for (let i = 1; i < values.length; i++) {
      if (String(values[i][1]) === householdId) {
        rows.push(values[i]);
      }
    }

    if (rows.length === 0) return null;

    const first = rows[0];
    const guests = rows.map(function(row) {
      return {
        guestNumber: row[5],
        firstName: row[6],
        lastName: row[7],
        status: row[8],
        type: row[9],
        meal: row[10],
        rehearsalDinner: row[11],
        morningAfterBrunch: row[12],
        comments: row[13]
      };
    });

    return {
      submittedAt: first[0],
      householdId: first[1],
      lookupFirstName: first[2],
      lookupLastName: first[3],
      guestCount: guests.length,
      guests: guests
    };
  }).filter(function(match) {
    return match !== null;
  });
}

function removeExistingRowsByHousehold_(sheet, householdId) {
  const values = sheet.getDataRange().getValues();

  for (let i = values.length - 1; i >= 1; i--) {
    if (String(values[i][1]) === String(householdId)) {
      sheet.deleteRow(i + 1);
    }
  }
}
