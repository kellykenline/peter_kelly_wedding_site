const SPREADSHEET_ID = "1hL4kQoY7HDOlotGPbBkvnGJj_lC54gg4AmrkjsHnC2Y";
const SHEET_NAME = "RSVP Responses";

const MAX_GUESTS = 8;
const MAX_NAME_LENGTH = 60;
const MAX_COMMENT_LENGTH = 1000;
const MIN_FORM_TIME_MS = 3000;
const MAX_FORM_AGE_MS = 24 * 60 * 60 * 1000;
const MAX_SUBMISSIONS_PER_HOUR = 10;

const VALID_STATUSES = ["Joyfully Accept", "Regretfully Decline"];
const VALID_GUEST_TYPES = ["Adult", "Child"];
const VALID_ADULT_MEALS = ["Filet", "Chicken", "Vegetarian"];
const VALID_CHILD_MEALS = ["Chicken", "Pasta"];
const VALID_YES_NO = ["Yes", "No"];

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
  const action = String((e && e.parameter && e.parameter.action) || "");
  const callback = safeCallback_((e && e.parameter && e.parameter.callback) || "callback");

  try {
    if (action === "lookup") {
      const firstName = validateName_(e.parameter.firstName, "First name");
      const lastName = validateName_(e.parameter.lastName, "Last name");
      const matches = findRsvpsByGuest_(firstName, lastName);

      return jsonp_(callback, {
        result: "success",
        found: matches.length > 0,
        matches: matches,
        data: matches.length === 1 ? matches[0] : null
      });
    }

    if (action === "verify") {
      const householdId = validateHouseholdId_(e.parameter.householdId);
      const submittedAt = validateText_(e.parameter.submittedAt, "Submitted timestamp", 50, false);
      const guestCount = validateInteger_(e.parameter.guestCount, "Guest count", 1, MAX_GUESTS);

      return jsonp_(callback, {
        result: "success",
        verified: verifyHouseholdWrite_(householdId, submittedAt, guestCount)
      });
    }

    return ContentService.createTextOutput("RSVP endpoint is active.");
  } catch (error) {
    return jsonp_(callback, {
      result: "error",
      message: error.message
    });
  }
}

function doPost(e) {
  const lock = LockService.getScriptLock();

  try {
    if (!e || !e.parameter || !e.parameter.payload) {
      throw new Error("Missing RSVP payload.");
    }

    const data = JSON.parse(e.parameter.payload);
    const validated = validateSubmission_(data);

    enforceRateLimit_(validated.lookupFirstName, validated.lookupLastName);

    lock.waitLock(15000);

    const ss = SpreadsheetApp.openById(SPREADSHEET_ID);
    const sheet = getOrCreateSheet_(ss);

    removeExistingRowsByHousehold_(sheet, validated.householdId);

    const rows = validated.guests.map(function(guest, index) {
      return [
        safeForSheet_(validated.submittedAt),
        safeForSheet_(validated.householdId),
        safeForSheet_(validated.lookupFirstName),
        safeForSheet_(validated.lookupLastName),
        validated.guests.length,
        index + 1,
        safeForSheet_(guest.firstName),
        safeForSheet_(guest.lastName),
        guest.status,
        guest.type,
        guest.meal,
        guest.rehearsalDinner,
        guest.morningAfterBrunch,
        safeForSheet_(guest.comments)
      ];
    });

    sheet.getRange(sheet.getLastRow() + 1, 1, rows.length, HEADERS.length).setValues(rows);
    SpreadsheetApp.flush();

    return json_({
      result: "success",
      householdId: validated.householdId,
      guestCount: rows.length
    });
  } catch (error) {
    return json_({
      result: "error",
      message: error.message
    });
  } finally {
    try {
      lock.releaseLock();
    } catch (ignored) {}
  }
}

function validateSubmission_(data) {
  if (!data || typeof data !== "object") {
    throw new Error("Invalid RSVP submission.");
  }

  if (String(data.honeypot || "").trim()) {
    throw new Error("Automated submission rejected.");
  }

  const now = Date.now();
  const openedAt = Number(data.formOpenedAt);
  if (!Number.isFinite(openedAt)) {
    throw new Error("Invalid form timestamp.");
  }

  const elapsed = now - openedAt;
  if (elapsed < MIN_FORM_TIME_MS) {
    throw new Error("The RSVP was submitted too quickly.");
  }
  if (elapsed > MAX_FORM_AGE_MS) {
    throw new Error("The RSVP form has expired. Please refresh the page and try again.");
  }

  const guests = Array.isArray(data.guests) ? data.guests : [];
  if (guests.length < 1 || guests.length > MAX_GUESTS) {
    throw new Error("The RSVP must contain between 1 and " + MAX_GUESTS + " guests.");
  }

  const submittedAt = validateIsoDate_(data.submittedAt);
  const householdId = data.householdId
    ? validateHouseholdId_(data.householdId)
    : Utilities.getUuid();

  const lookupFirstName = validateName_(data.lookupFirstName, "Lookup first name");
  const lookupLastName = validateName_(data.lookupLastName, "Lookup last name");

  const validatedGuests = guests.map(function(guest, index) {
    return validateGuest_(guest, index + 1);
  });

  return {
    submittedAt: submittedAt,
    householdId: householdId,
    lookupFirstName: lookupFirstName,
    lookupLastName: lookupLastName,
    guests: validatedGuests
  };
}

function validateGuest_(guest, guestNumber) {
  if (!guest || typeof guest !== "object") {
    throw new Error("Guest " + guestNumber + " is invalid.");
  }

  const firstName = validateName_(guest.firstName, "Guest " + guestNumber + " first name");
  const lastName = validateName_(guest.lastName, "Guest " + guestNumber + " last name");
  const status = validateChoice_(guest.status, VALID_STATUSES, "Guest " + guestNumber + " attendance");
  const comments = validateText_(guest.comments || "", "Guest " + guestNumber + " comments", MAX_COMMENT_LENGTH, true);

  let type = "";
  let meal = "";
  let rehearsalDinner = "No";
  let morningAfterBrunch = "No";

  if (status === "Joyfully Accept") {
    type = validateChoice_(guest.type, VALID_GUEST_TYPES, "Guest " + guestNumber + " type");
    const validMeals = type === "Child" ? VALID_CHILD_MEALS : VALID_ADULT_MEALS;
    meal = validateChoice_(guest.meal, validMeals, "Guest " + guestNumber + " meal");
    rehearsalDinner = validateChoice_(guest.rehearsalDinner || "No", VALID_YES_NO, "Guest " + guestNumber + " rehearsal dinner");
    morningAfterBrunch = validateChoice_(guest.morningAfterBrunch || "No", VALID_YES_NO, "Guest " + guestNumber + " brunch");
  }

  return {
    firstName: firstName,
    lastName: lastName,
    status: status,
    type: type,
    meal: meal,
    rehearsalDinner: rehearsalDinner,
    morningAfterBrunch: morningAfterBrunch,
    comments: comments
  };
}

function validateName_(value, label) {
  const text = validateText_(value, label, MAX_NAME_LENGTH, false);
  if (/[<>{}\[\]\\]/.test(text)) {
    throw new Error(label + " contains unsupported characters.");
  }
  return text;
}

function validateText_(value, label, maxLength, allowEmpty) {
  let text = String(value == null ? "" : value);
  text = text.replace(/[\u0000-\u0008\u000B\u000C\u000E-\u001F\u007F]/g, "").trim();

  if (!allowEmpty && !text) {
    throw new Error(label + " is required.");
  }
  if (text.length > maxLength) {
    throw new Error(label + " is too long.");
  }

  return text;
}

function validateChoice_(value, allowed, label) {
  const text = String(value == null ? "" : value);
  if (allowed.indexOf(text) === -1) {
    throw new Error(label + " is invalid.");
  }
  return text;
}

function validateInteger_(value, label, min, max) {
  const number = Number(value);
  if (!Number.isInteger(number) || number < min || number > max) {
    throw new Error(label + " is invalid.");
  }
  return number;
}

function validateIsoDate_(value) {
  const text = validateText_(value, "Submitted timestamp", 50, false);
  const time = Date.parse(text);
  if (!Number.isFinite(time)) {
    throw new Error("Submitted timestamp is invalid.");
  }
  return new Date(time).toISOString();
}

function validateHouseholdId_(value) {
  const text = validateText_(value, "Household ID", 100, false);
  if (!/^[A-Za-z0-9_-]{8,100}$/.test(text)) {
    throw new Error("Household ID is invalid.");
  }
  return text;
}

function safeForSheet_(value) {
  const text = String(value == null ? "" : value);
  return /^[=+\-@]/.test(text) ? "'" + text : text;
}

function safeCallback_(value) {
  const callback = String(value || "callback");
  return /^[A-Za-z_$][A-Za-z0-9_$.]{0,100}$/.test(callback) ? callback : "callback";
}

function enforceRateLimit_(firstName, lastName) {
  const cache = CacheService.getScriptCache();
  const key = "rsvp_rate_" + normalize_(firstName + "_" + lastName).replace(/[^a-z0-9_]/g, "").slice(0, 80);
  const current = Number(cache.get(key) || "0");

  if (current >= MAX_SUBMISSIONS_PER_HOUR) {
    throw new Error("Too many RSVP submissions. Please wait and try again later.");
  }

  cache.put(key, String(current + 1), 3600);
}

function json_(value) {
  return ContentService
    .createTextOutput(JSON.stringify(value))
    .setMimeType(ContentService.MimeType.JSON);
}

function jsonp_(callback, value) {
  return ContentService
    .createTextOutput(callback + "(" + JSON.stringify(value) + ");")
    .setMimeType(ContentService.MimeType.JAVASCRIPT);
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
    const rows = values.filter(function(row, index) {
      return index > 0 && String(row[1]) === householdId;
    });

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

function verifyHouseholdWrite_(householdId, submittedAt, guestCount) {
  const ss = SpreadsheetApp.openById(SPREADSHEET_ID);
  const sheet = getOrCreateSheet_(ss);
  const values = sheet.getDataRange().getValues();

  const matchingRows = values.filter(function(row, index) {
    return index > 0 &&
      String(row[1]) === String(householdId) &&
      String(row[0]) === String(submittedAt);
  });

  return matchingRows.length === guestCount;
}

function removeExistingRowsByHousehold_(sheet, householdId) {
  const values = sheet.getDataRange().getValues();

  for (let i = values.length - 1; i >= 1; i--) {
    if (String(values[i][1]) === String(householdId)) {
      sheet.deleteRow(i + 1);
    }
  }
}
