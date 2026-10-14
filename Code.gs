

New Template


/**

 * VEHICLE INSURANCE RENEWAL WHATSAPP REMINDER

 * Google Apps Script + Google Sheets + Pinbot API

 *

 * Required Script Properties:

 *   PINBOT_API_KEY : 

 *   PINBOT_PHONE_NUMBER_ID :

 *

 * Daily reminder time: approximately 8:00 AM Asia/Kolkata.

 * Live mode: messages go to each eligible row's actual mobile number.

 */


const CONFIG = {

  SPREADSHEET_ID: ' "Enter the Google Sheet Id" ',

  SHEET_NAME: 'Sheet1',

  TEMPLATE_NAME: ' - ',

  LANGUAGE_CODE: 'en',

  HEADER_IMAGE_URL: 'https://whatsappdata.s3.ap-south-1.amazonaws.com/userMedia/673c5829e2a273388ffdc550ef94491c/testing.jpeg',

  TIME_ZONE: 'Asia/Kolkata',

  REMINDER_DAYS: [10, 8, 3, 2, 1, 0],

  LOG_SHEET_NAME: 'Insurance_Reminder_Logs',

  TEST_MODE: false,

  TEST_TO_NUMBER: '9621192312',

  PHONE_HEADER: 'Mobile number as per our data',

  OWNER_HEADER: 'owner_name',

  REG_HEADER: 'reg_no',

  EXPIRY_HEADER: 'vehicle_insurance_upto'

};



/**

 * Checks sheet access and creates the log sheet if needed.

 * Does NOT send WhatsApp messages.

 */

function testSheetAccess() {

  const ss = SpreadsheetApp.openById(CONFIG.SPREADSHEET_ID);

  const sheet = ss.getSheetByName(CONFIG.SHEET_NAME);

  if (!sheet) {

    throw new Error('Sheet tab not found: "' + CONFIG.SHEET_NAME + '".');

  }

  Logger.log('Spreadsheet: ' + ss.getName());

  Logger.log('Sheet: ' + sheet.getName());

  Logger.log('Data rows: ' + Math.max(0, sheet.getLastRow() - 1));

  getOrCreateLogSheet_(ss);

  Logger.log('Log sheet ready. No WhatsApp message was sent.');

}



/**

 * Main reminder sender.

 * Sends only on 10, 8, 3, 2, 1 or 0 days before/on expiry.

 * Successful registration + expiry + days-left combinations are not sent again.

 */

function runInsuranceReminders() {

  const lock = LockService.getScriptLock();

  if (!lock.tryLock(30000)) {

    throw new Error('Another reminder run is already in progress.');

  }



  try {

    const props = PropertiesService.getScriptProperties();

    const apiKey = props.getProperty('PINBOT_API_KEY');

    const phoneNumberId = props.getProperty('PINBOT_PHONE_NUMBER_ID');



    if (!apiKey || !phoneNumberId) {

      throw new Error('Add PINBOT_API_KEY and PINBOT_PHONE_NUMBER_ID in Script Properties.');

    }

    if (CONFIG.TEST_MODE && !normalizePhone_(CONFIG.TEST_TO_NUMBER)) {

      throw new Error('TEST_MODE is true. Enter a valid TEST_TO_NUMBER.');

    }



    const ss = SpreadsheetApp.openById(CONFIG.SPREADSHEET_ID);

    const sheet = ss.getSheetByName(CONFIG.SHEET_NAME);

    if (!sheet) throw new Error('Sheet tab not found: ' + CONFIG.SHEET_NAME);



    const values = sheet.getDataRange().getValues();

    if (values.length < 2) {

      Logger.log('No vehicle data rows found.');

      return;

    }



    const headers = values[0].map(function (v) { return String(v).trim(); });

    const phoneCol = findHeader_(headers, CONFIG.PHONE_HEADER);

    const ownerCol = findHeader_(headers, CONFIG.OWNER_HEADER);

    const regCol = findHeader_(headers, CONFIG.REG_HEADER);

    const expiryCols = findAllHeaders_(headers, CONFIG.EXPIRY_HEADER);



    if (phoneCol < 0 || ownerCol < 0 || regCol < 0 || expiryCols.length === 0) {

      throw new Error('Required headers not found: ' +

        [CONFIG.PHONE_HEADER, CONFIG.OWNER_HEADER, CONFIG.REG_HEADER,

          CONFIG.EXPIRY_HEADER].join(' | '));

    }



    const logSheet = getOrCreateLogSheet_(ss);

    const sentKeys = getExistingSuccessfulKeys_(logSheet);

    const today = dateOnlyInTimeZone_(new Date(), CONFIG.TIME_ZONE);

    let dueCount = 0;

    let successCount = 0;

    let failedCount = 0;



    for (let r = 1; r < values.length; r++) {

      const row = values[r];

      let regNo = '';

      let rawPhone = '';



      try {

        const owner = String(row[ownerCol] || '').trim();

        regNo = String(row[regCol] || '').trim();

        rawPhone = row[phoneCol];



        const customerPhone = normalizePhone_(rawPhone);

        const expiry = getExpiryDate_(row, expiryCols);



        if (!regNo || !customerPhone || !expiry) {

          log_(logSheet, new Date(), regNo, String(rawPhone || ''), '',

            'SKIPPED',

            'Missing registration number, valid mobile number, or valid insurance expiry date.');

          continue;

        }



        const expiryDay = dateOnlyInTimeZone_(expiry, CONFIG.TIME_ZONE);

        const daysLeft = Math.round((expiryDay.getTime() - today.getTime()) / 86400000);

        if (CONFIG.REMINDER_DAYS.indexOf(daysLeft) === -1) continue;



        dueCount++;

        const dedupeKey = makeKey_(regNo, expiryDay, daysLeft);

        if (sentKeys.has(dedupeKey)) continue;



        const recipient = CONFIG.TEST_MODE

          ? normalizePhone_(CONFIG.TEST_TO_NUMBER)

          : customerPhone;



        const bodyValues = [

          owner || 'Customer',

          regNo,

          Utilities.formatDate(expiryDay, CONFIG.TIME_ZONE, 'dd-MM-yyyy')

        ];



        const result = sendPinbotTemplate_(apiKey, phoneNumberId, recipient, bodyValues);

        const success = result.httpCode >= 200 && result.httpCode < 300;



        log_(logSheet, new Date(), regNo, recipient, dedupeKey,

          success ? 'SENT' : 'FAILED',

          'Days until expiry: ' + daysLeft + ' | HTTP ' + result.httpCode + ' | ' + result.body);



        if (success) {

          sentKeys.add(dedupeKey);

          successCount++;

        } else {

          failedCount++;

        }

      } catch (err) {

        failedCount++;

        log_(logSheet, new Date(), regNo, String(rawPhone || ''), '',

          'ERROR', String(err && err.stack ? err.stack : err));

      }

    }



    Logger.log('Run finished. Due rows: ' + dueCount +

      ', API accepted: ' + successCount +

      ', failed/errors: ' + failedCount +

      ', TEST_MODE: ' + CONFIG.TEST_MODE);

  } finally {

    lock.releaseLock();

  }

}



/**

 * Sends the approved template with IMAGE header and three BODY parameters.

 */

function sendPinbotTemplate_(apiKey, phoneNumberId, recipient, bodyValues) {

  const url = 'https://partnersv1.pinbot.ai/v3/' +

    encodeURIComponent(phoneNumberId) + '/messages';



  const payload = {

    messaging_product: 'whatsapp',

    recipient_type: 'individual',

    to: recipient,

    type: 'template',

    template: {

      name: CONFIG.TEMPLATE_NAME,

      language: { code: CONFIG.LANGUAGE_CODE },

      components: [

        {

          type: 'header',

          parameters: [

            {

              type: 'image',

              image: { link: CONFIG.HEADER_IMAGE_URL }

            }

          ]

        },

        {

          type: 'body',

          parameters: bodyValues.map(function (value) {

            return { type: 'text', text: String(value) };

          })

        }

      ]

    }

  };



  const response = UrlFetchApp.fetch(url, {

    method: 'post',

    contentType: 'application/json',

    headers: { apikey: apiKey },

    payload: JSON.stringify(payload),

    muteHttpExceptions: true

  });



  return {

    httpCode: response.getResponseCode(),

    body: response.getContentText()

  };

}



/**

 * Creates/replaces the daily trigger for 8 AM IST.

 * Run this function manually ONCE after saving the code.

 */

function setupDailyTrigger() {

  ScriptApp.getProjectTriggers()

    .filter(function (trigger) {

      return trigger.getHandlerFunction() === 'runInsuranceReminders';

    })

    .forEach(function (trigger) {

      ScriptApp.deleteTrigger(trigger);

    });



  ScriptApp.newTrigger('runInsuranceReminders')

    .timeBased()

    .everyDays(1)

    .atHour(8)

    .inTimezone(CONFIG.TIME_ZONE)

    .create();



  Logger.log('Daily insurance reminders scheduled for approximately 8 AM IST.');

}



/** Finds a single exact header name, ignoring case. */

function findHeader_(headers, wanted) {

  return headers.findIndex(function (header) {

    return header.toLowerCase() === wanted.toLowerCase();

  });

}



/** Finds every duplicate expiry header, if the sheet has more than one. */

function findAllHeaders_(headers, wanted) {

  const matches = [];

  headers.forEach(function (header, index) {

    if (header.toLowerCase() === wanted.toLowerCase()) matches.push(index);

  });

  return matches;

}



/** Uses the first expiry column containing a valid date. */

function getExpiryDate_(row, expiryCols) {

  for (let i = 0; i < expiryCols.length; i++) {

    const parsed = parseDate_(row[expiryCols[i]]);

    if (parsed) return parsed;

  }

  return null;

}



/** Supports Date cells, dd-MM-yyyy, dd/MM/yyyy, and yyyy-MM-dd. */

function parseDate_(value) {

  if (value instanceof Date && !isNaN(value.getTime())) {

    return new Date(value.getFullYear(), value.getMonth(), value.getDate());

  }

  if (value === null || value === undefined || String(value).trim() === '') return null;



  const s = String(value).trim();

  let match = s.match(/^(\d{1,2})[-/](\d{1,2})[-/](\d{4})$/);



  if (match) {

    const day = Number(match[1]);

    const month = Number(match[2]);

    const year = Number(match[3]);

    const date = new Date(year, month - 1, day);

    if (date.getFullYear() === year &&

        date.getMonth() === month - 1 &&

        date.getDate() === day) return date;

    return null;

  }



  match = s.match(/^(\d{4})-(\d{1,2})-(\d{1,2})$/);

  if (match) {

    const year = Number(match[1]);

    const month = Number(match[2]);

    const day = Number(match[3]);

    const date = new Date(year, month - 1, day);

    if (date.getFullYear() === year &&

        date.getMonth() === month - 1 &&

        date.getDate() === day) return date;

  }

  return null;

}



/** Converts a date to date-only in the configured timezone. */

function dateOnlyInTimeZone_(date, timeZone) {

  const year = Number(Utilities.formatDate(date, timeZone, 'yyyy'));

  const month = Number(Utilities.formatDate(date, timeZone, 'MM'));

  const day = Number(Utilities.formatDate(date, timeZone, 'dd'));

  return new Date(year, month - 1, day);

}



/** Normalizes Indian 10-digit numbers to country code 91. */

function normalizePhone_(value) {

  if (value === null || value === undefined) return '';

  let digits = String(value).trim().replace(/[^\d]/g, '');

  if (!digits) return '';

  if (digits.length === 10) digits = '91' + digits;

  if (digits.length < 11 || digits.length > 15) return '';

  return digits;

}



/** Unique key used to avoid repeating a successful reminder for the same day/key. */

function makeKey_(regNo, expiry, daysLeft) {

  const expiryString = Utilities.formatDate(expiry, CONFIG.TIME_ZONE, 'yyyy-MM-dd');

  return [String(regNo).trim().toUpperCase(), expiryString, daysLeft].join('|');

}



/** Creates the log sheet if it does not exist. */

function getOrCreateLogSheet_(ss) {

  let sheet = ss.getSheetByName(CONFIG.LOG_SHEET_NAME);

  if (!sheet) {

    sheet = ss.insertSheet(CONFIG.LOG_SHEET_NAME);

    sheet.appendRow([

      'Timestamp', 'Registration No', 'Recipient',

      'Deduplication Key', 'Status', 'Details'

    ]);

    sheet.setFrozenRows(1);

  }

  return sheet;

}



/** Loads successful deduplication keys from the log sheet. */

function getExistingSuccessfulKeys_(logSheet) {

  const data = logSheet.getDataRange().getValues();

  const keys = new Set();

  for (let i = 1; i < data.length; i++) {

    const key = String(data[i][3] || '');

    const status = String(data[i][4] || '');

    if (key && status === 'SENT') keys.add(key);

  }

  return keys;

}

/** Appends one result to the log sheet. */

function log_(logSheet, timestamp, regNo, recipient, key, status, details) {

  logSheet.appendRow([timestamp, regNo, recipient, key, status, details]);

}


/*
SETUP:
1. Add PINBOT_API_KEY and PINBOT_PHONE_NUMBER_ID under Project Settings > Script properties.
2. Save the project.
3. Run testSheetAccess() once to check the spreadsheet and create the log sheet. This does not send messages.
4. Run setupDailyTrigger() once to create/replace the daily trigger (approximately 8 AM Asia/Kolkata).
5. Since TEST_MODE is false, runInsuranceReminders() and the daily trigger send to eligible customers' actual mobile numbers.
*/


