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
    const data = findRsvpByGuest_(e.parameter.firstName || "", e.parameter.lastName || "");
    const output = JSON.stringify({ result: "success", found: data !== null, data: data });
    return ContentService.createTextOutput(callback + "(" + output + ");").setMimeType(ContentService.MimeType.JAVASCRIPT);
  }

  return ContentService.createTextOutput("RSVP endpoint is active.");
}

function doPost(e) {
  try {
    const data = JSON.parse(e.parameter.payload);
    const ss = SpreadsheetApp.openById(SPREADSHEET_ID);
    const sheet = getOrCreateSheet_(ss);
    const householdId = data.householdId || Utilities.getUuid();

    removeExistingRowsByHousehold_(sheet, householdId);

    data.guests.forEach(function(guest) {
      sheet.appendRow([
        data.submittedAt,
        householdId,
        data.lookupFirstName,
        data.lookupLastName,
        data.guestCount,
        guest.guestNumber,
        guest.firstName,
        guest.lastName,
        guest.status,
        guest.type,
        guest.meal,
        guest.rehearsalDinner,
        guest.morningAfterBrunch,
        guest.comments
      ]);
    });

    return ContentService.createTextOutput("success");
  } catch (error) {
    return ContentService.createTextOutput("error: " + error.message);
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

function findRsvpByGuest_(firstName, lastName) {
  const ss = SpreadsheetApp.openById(SPREADSHEET_ID);
  const sheet = getOrCreateSheet_(ss);
  const values = sheet.getDataRange().getValues();

  if (values.length <= 1) return null;

  const targetFirst = normalize_(firstName);
  const targetLast = normalize_(lastName);
  let householdId = null;

  for (let i = 1; i < values.length; i++) {
    if (normalize_(values[i][6]) === targetFirst && normalize_(values[i][7]) === targetLast) {
      householdId = values[i][1];
      break;
    }
  }

  if (!householdId) return null;

  const rows = [];
  for (let i = 1; i < values.length; i++) {
    if (String(values[i][1]) === String(householdId)) rows.push(values[i]);
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
    guestCount: first[4],
    guests: guests
  };
}

function removeExistingRowsByHousehold_(sheet, householdId) {
  const values = sheet.getDataRange().getValues();

  for (let i = values.length - 1; i >= 1; i--) {
    if (String(values[i][1]) === String(householdId)) {
      sheet.deleteRow(i + 1);
    }
  }
}
