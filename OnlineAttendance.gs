/**
 * ONLINE Service Attendance – Confirmation Sender (Option B: date-only matching)
 * Sheets:
 * - Schedule: A=Event ID, B=Event Name, C=Date, D=Time, E=Pastor, F=Link, G=Notes
 * - Email Content: C2=Subject, C3=Body(HTML), C4=ImageURL(optional), C5=From Name
 */

const ss = SpreadsheetApp.getActiveSpreadsheet();
const SCHEDULE_SHEET = 'Schedule';
const EMAIL_SHEET = 'Email Content';
const LOGS_SHEET = 'Logs'; // <— added

// === MAIN HANDLER (Installable Trigger: From spreadsheet -> On form submit) ===
// MODIFIED to pass the email to the logger
function onFormSubmit(e) {
  try {
    const nv = (e && e.namedValues) ? e.namedValues : {};
    const email   = getFirst(nv['Email Address']);
    const first   = getFirst(nv['First Name']);
    const last    = getFirst(nv['Last Name']);
    const choice  = getFirst(nv['Pick a Sunday Service you will attend']); // free text date

    if (!email || !choice) {
      // log not sent reason (kept your original behavior; just logging)
      appendLog({
        email: email || '', // <--- ADDED
        first: first || '',
        last: last || '',
        eventName: '',
        dateDisplay: choice || '',
        link: '',
        sentTs: '',
        status: !email ? 'not sent — missing recipient email' : 'not sent — missing date choice'
      });
      Logger.log('Missing email or selected date. Aborting.');
      return;
    }

    // Find matching row in Schedule by date text (robust normalization)
    const match = findScheduleRowByDateText(choice);
    if (!match) {
      // Graceful fallback if we cannot match the date
      const fallback = getEmailContent();
      const map = {
        FirstName: first || '',
        LastName: last || '',
        EventName: '(TBD)',
        Date: choice || '',
        Time: '',
        Pastor: '',
        Notes: 'We could not locate the event date you selected. Our team will follow up with the correct link shortly.',
        Link: ''
      };
      const subject = replaceTpl(fallback.subject || 'Your registration', map);
      let body = replaceTpl(fallback.body || '<p>Thank you — we will send the link soon.</p>', map);
      if (fallback.imageUrl) {
        body = `<p><img src="${fallback.imageUrl}" style="max-width:100%;height:auto;"></p>` + body;
      }
      try {
        MailApp.sendEmail({ to: email, subject, htmlBody: body, name: (fallback.fromName || undefined) });
        appendLog({
          email: email, // <--- ADDED
          first: first || '',
          last: last || '',
          eventName: '(TBD)',
          dateDisplay: choice || '',
          link: '',
          sentTs: new Date(),
          status: 'sent'
        });
        Logger.log('No schedule match. Sent fallback notice.');
      } catch (sendErr) {
        appendLog({
          email: email, // <--- ADDED
          first: first || '',
          last: last || '',
          eventName: '(TBD)',
          dateDisplay: choice || '',
          link: '',
          sentTs: '',
          status: 'not sent — send error: ' + String(sendErr && sendErr.message || sendErr)
        });
        Logger.log('Send error (fallback): ' + sendErr);
      }
      return;
    }

    const content = getEmailContent();

    // Build placeholder map
    const map = {
      FirstName: first || '',
      LastName: last || '',
      EventName: match.eventName || '',
      Date: match.datePretty || '',
      Time: match.time || '',
      Pastor: match.pastor || '',
      Notes: match.notes || '',
      Link: match.link || ''
    };

    const subject = replaceTpl(content.subject || 'Your registration for {{EventName}}', map);
    let body = replaceTpl(content.body || '<p>Thank you for registering.</p>', map);
    if (content.imageUrl) {
      body = `<p><img src="${content.imageUrl}" style="max-width:100%;height:auto;"></p>` + body;
    }

    try {
      MailApp.sendEmail({
        to: email,
        subject: subject,
        htmlBody: body,
        name: (content.fromName ? String(content.fromName).trim() : undefined) // Email Content!C5
      });

      // log success
      appendLog({
        email: email, // <--- ADDED
        first: first || '',
        last: last || '',
        eventName: match.eventName || '',
        dateDisplay: match.datePretty || '',
        link: match.link || '',
        sentTs: new Date(),
        status: 'sent'
      });

      Logger.log(`Sent to ${email} for ${match.datePretty}`);
    } catch (sendErr) {
      appendLog({
        email: email, // <--- ADDED
        first: first || '',
        last: last || '',
        eventName: match.eventName || '',
        dateDisplay: match.datePretty || '',
        link: match.link || '',
        sentTs: '',
        status: 'not sent — send error: ' + String(sendErr && sendErr.message || sendErr)
      });
      Logger.log('Send error: ' + sendErr);
    }
  } catch (err) {
    Logger.log('onFormSubmit error: ' + err + '\n' + (err.stack || ''));
  }
}

// === LOOKUP: match free-text date to Schedule!C ===
function findScheduleRowByDateText(inputText) {
  const sh = ss.getSheetByName(SCHEDULE_SHEET);
  if (!sh) return null;

  const tz = ss.getSpreadsheetTimeZone() || 'America/New_York';
  const values = sh.getDataRange().getValues();
  if (values.length < 2) return null;

  const header = values[0];
  // Expecting A=Event ID, B=Event Name, C=Date, D=Time, E=Pastor, F=Link, G=Notes

  const normInput = normalizeDateText(inputText);

  // Build candidate map: normalized text -> array of row indices
  const candidates = [];
  for (let r = 1; r < values.length; r++) {
    const row = values[r];
    const dateCell = row[2]; // col C
    if (!(dateCell instanceof Date)) continue;

    const formats = [
      "MMMM d, yyyy", "MMM d, yyyy",
      "M/d/yyyy", "MM/dd/yyyy",
      "d MMM yyyy", "d MMMM yyyy",
      "yyyy-MM-dd"
    ];

    const normSet = new Set();
    formats.forEach(fmt => {
      const str = Utilities.formatDate(dateCell, tz, fmt);
      normSet.add(normalizeDateText(str));
    });

    // Keep row if any candidate equals normalized input
    if ([...normSet].some(k => k === normInput)) {
      const eventName = row[1] || '';
      const time      = row[3] || '';
      const pastor    = row[4] || '';
      const link      = row[5] || '';
      const notes     = row[6] || '';

      // Pretty date for email output
      const datePretty = Utilities.formatDate(dateCell, tz, "MMMM d, yyyy");
      return { rowIndex: r+1, eventName, time, pastor, link, notes, datePretty, dateObj: dateCell };
    }
  }
  return null;
}

// === Read Email Content (Subject/Body/Image/From Name) ===
function getEmailContent() {
  const sh = ss.getSheetByName(EMAIL_SHEET);
  if (!sh) return { subject: '', body: '', imageUrl: '', fromName: '' };
  const subject  = sh.getRange('C2').getValue(); // Subject
  const body     = sh.getRange('C3').getValue(); // Body (HTML)
  const imageUrl = sh.getRange('C4').getValue(); // Optional Image URL
  const fromName = sh.getRange('C5').getValue(); // From Name (customizable)
  return { subject, body, imageUrl, fromName };
}

function appendLog({ email, first, last, eventName, dateDisplay, link, sentTs, status }) {
  const sh = getOrCreateLogsSheet();
  // Ensure header exists
  if (sh.getLastRow() === 0) {
    sh.appendRow(['Email Address', 'First Name', 'Last Name', 'Event Name', 'Date', 'URL', 'Email Sent Timestamp', 'Status']);
  }
  sh.appendRow([
    email || '', // <-- ADDED
    first || '',
    last || '',
    eventName || '',
    dateDisplay || '',
    link || '',
    sentTs || '',
    status || ''
  ]);
}

function getOrCreateLogsSheet() {
  let sh = ss.getSheetByName(LOGS_SHEET);
  if (!sh) {
    sh = ss.insertSheet(LOGS_SHEET);
    sh.appendRow(['First Name', 'Last Name', 'Event Name', 'Date', 'URL', 'Email Sent Timestamp', 'Status']);
  }
  return sh;
}

// === Helpers ===
function getFirst(v) {
  if (v == null) return '';
  if (Array.isArray(v)) return String(v[0] || '');
  return Array.isArray(v.values) ? String(v.values[0] || '') : String(v);
}

function replaceTpl(str, map) {
  if (!str) return '';
  return String(str).replace(/{{\s*(\w+)\s*}}/g, (m, k) => (map[k] != null ? String(map[k]) : ''));
}

function normalizeDateText(s) {
  return String(s || '')
    .toLowerCase()
    .replace(/(\d+)(st|nd|rd|th)\b/g, '$1') // 1st -> 1
    .replace(/[,\u00A0]/g, ' ')              // commas & nbsp -> space
    .replace(/\s+/g, ' ')                    // collapse spaces
    .trim();
}

/**
 * Optional: manual test helper
 * - Fill these with sample data then run sendTestOnce() to verify templating & sending.
 */
function sendTestOnce() {
  const fake = {
    namedValues: {
      'Email Address': ['eduardo.purewaterautomations@gmail.com'],
      'First Name': ['Test'],
      'Last Name': ['User'],
      'Pick a Sunday Service you will attend': ['August 24, 2025'] // must exist in Schedule!C
    }
  };
  onFormSubmit(fake);
}
/**
 * Sends reminder emails for any event scheduled for the current day.
 * This function should be run on a daily time-based trigger.
 */
function sendDailyReminders() {
  const now = new Date();
  const tz = ss.getSpreadsheetTimeZone();
  const todayDateText = Utilities.formatDate(now, tz, "MMMM d, yyyy");
  
  Logger.log(`Trigger fired. Running reminder check for today: ${todayDateText}`);

  const eventDetails = findScheduleRowByDateText(todayDateText);
  if (!eventDetails) {
    Logger.log(`No event found in 'Schedule' for today, ${todayDateText}. No reminders sent.`);
    return;
  }

  const emailSheet = ss.getSheetByName(EMAIL_SHEET);
  if (!emailSheet) {
    Logger.log(`'${EMAIL_SHEET}' sheet not found. Cannot send reminders.`);
    return;
  }
  const reminderSubjectTpl = emailSheet.getRange('C7').getValue();
  const reminderBodyTpl = emailSheet.getRange('C8').getValue();
  const imageUrl = emailSheet.getRange('C4').getValue();
  const fromName = emailSheet.getRange('C5').getValue();

  if (!reminderSubjectTpl || !reminderBodyTpl) {
    Logger.log("Reminder subject or body template is missing from 'Email Content' sheet (cells C7, C8).");
    return;
  }

  const logsSheet = ss.getSheetByName(LOGS_SHEET);
  if (!logsSheet || logsSheet.getLastRow() < 2) {
    Logger.log('No logs found. No reminders to send.');
    return;
  }
  
  const logsData = logsSheet.getDataRange().getValues();
  const headers = logsData.shift(); 
  const emailIndex = headers.indexOf('Email Address');
  const dateIndex = headers.indexOf('Date');
  const statusIndex = headers.indexOf('Status');
  const firstNameIndex = headers.indexOf('First Name');
  
  if (emailIndex === -1 || dateIndex === -1 || statusIndex === -1) {
      Logger.log("Could not find required columns ('Email Address', 'Date', 'Status') in Logs sheet.");
      return;
  }
  
  let remindersSent = 0;
  for (const row of logsData) {
    const logStatus = String(row[statusIndex]).trim();
    const rawLogDate = row[dateIndex];
    let formattedLogDate;

    if (rawLogDate instanceof Date && !isNaN(rawLogDate)) {
      formattedLogDate = Utilities.formatDate(rawLogDate, tz, "MMMM d, yyyy");
    } else {
      formattedLogDate = String(rawLogDate).trim();
    }
    
    if (logStatus === 'sent' && formattedLogDate === todayDateText) {
      const recipientEmail = row[emailIndex];
      const firstName = row[firstNameIndex];
      
      if (!recipientEmail) continue;

      const map = {
        FirstName: firstName || '',
        EventName: eventDetails.eventName || '',
        Date: eventDetails.datePretty || '',
        Time: eventDetails.time || '',
        Pastor: eventDetails.pastor || '',
        Notes: eventDetails.notes || '',
        Link: eventDetails.link || ''
      };
      
      const subject = replaceTpl(reminderSubjectTpl, map);
      let body = replaceTpl(reminderBodyTpl, map);
      if (imageUrl) {
        body = `<p><img src="${imageUrl}" style="max-width:100%;height:auto;"></p>` + body;
      }
      
      try {
        MailApp.sendEmail({
          to: recipientEmail,
          subject: subject,
          htmlBody: body,
          name: (fromName ? String(fromName).trim() : undefined)
        });
        Logger.log(`Reminder sent to ${recipientEmail} for ${todayDateText}`);
        remindersSent++;
      } catch(err) {
        Logger.log(`Failed to send reminder to ${recipientEmail}. Error: ${err.message}`);
      }
    }
  }
  
  Logger.log(`Reminder process finished. Sent ${remindersSent} emails.`);
}
