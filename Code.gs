/** @OnlyCurrentDoc */

/**
 * KUAS Dormitory Attendance Tracker
 * Google Apps Script bound to the attendance Google Sheet.
 * v7.0: NEW - Invite-Only events (small events where only specific residents
 *       are invited).
 *       - New menu item: Events -> Import Invite-Only Event (Tick Attendance),
 *         backed by a new dialog file ImportInviteOnlyEvent.html, plus a new
 *         section 4 on BulkActionTemplates (Invited ID | Dorm | Attended?
 *         with checkboxes) to fill in and copy-paste from.
 *       - The generated event tab lists every invited resident in column A
 *         with an "Attended?" checkbox in column B. Boxes can be ticked in
 *         the template before pasting, or later on the event tab itself
 *         (ticking a box auto-refreshes, same as typing an ID does).
 *       - Counting rule: an invite-only event counts toward a resident's
 *         Events Attended AND Eligible Events ONLY if their box is ticked.
 *         Invited no-shows and everyone who wasn't invited are untouched, so
 *         a small event can only ever raise someone's Attendance %, never
 *         drag down the whole building's numbers. Dorm scope (E6) is ignored
 *         for these events - the invite list is the scope.
 *       - Under the hood: EventLog gets a 7th column (G, "Event Type") and
 *         AttendanceLog a 10th (J, "Event Type"). Regular events leave it
 *         blank. The Eligible Events formula now skips invite-only events in
 *         EventLog and instead adds one per ticked invite-only attendance.
 *         Existing data needs no migration - blank = regular event.
 *       - An event tab is invite-only when B6 reads "Invite-Only" (A6 is
 *         labelled "Event Type"). Regular event tabs leave A6:B6 empty.
 *       - Data Audit also flags past invite-only events with nothing ticked.
 *
 * v6.9: PRIVACY - reduced the Google permissions this script asks for.
 *       - Added the @OnlyCurrentDoc annotation (first line of this file).
 *         The script now only gets access to THIS spreadsheet, instead of
 *         "See, edit, create, and delete all your Google Sheets
 *         spreadsheets". Nothing in this project ever opened another file,
 *         so nothing is lost.
 *       - Replaced the two installable triggers (handleAttendanceEdit /
 *         handleSpreadsheetChange) with Google's built-in simple onEdit()
 *         trigger. This removes the "Allow this application to run when you
 *         are not present" permission entirely.
 *         * Auto-refresh on typing/pasting attendee IDs still works (onEdit).
 *         * Deleted event tabs are still cleaned out of EventLog /
 *           AttendanceLog: every refresh rebuilds both logs from the tabs
 *           that still exist, onOpen now also runs the orphan cleanup, and
 *           Events -> Clean Up Deleted Event References still works.
 *       - The only other remaining permission is "Display and run third-party
 *         web content in prompts and sidebars" - that is Google's wording for
 *         the script showing its OWN prompts/alerts and the Bulk Register /
 *         Bulk Deactivate / Import Event dialogs. It can't be removed without
 *         removing the menus themselves.
 *
 * v6.8: Per-dorm table now has a blank spacer column at E, with "Eligible
 *       Event Slots" moved to F - matches the label / value / blank-spacer /
 *       label / value pattern already used in the row above it (D16/E16,
 *       blank F16, G16/H16).
 *
 * v6.7: BUGFIX - Active Students / Events Attended / Eligible Event Slots
 *       (and the top-row counts: Total Events This Year, Total Attendance
 *       Records, Total Active Students, Dorms Tracked) were inheriting a
 *       "0.00%" format from somewhere after clearFormat() ran, turning a
 *       plain count like 4 into "400.00%". clearFormat() resets a cell to
 *       "no override," which isn't the same as guaranteeing a specific
 *       format. Every numeric cell in the Current Summary block now gets its
 *       number format set explicitly and unconditionally right after being
 *       written, so nothing can be silently inherited again.
 *
 * v6.6: BUGFIX - breakApart() only undoes a merge; it doesn't clear whatever
 *       VALUE was sitting in a cell that had been hidden underneath that
 *       merge (e.g. a second "Eligible Event Slots" label typed into F20 at
 *       some point in the past, invisible the whole time it lived under a
 *       wider merged header). Unmerging in v6.5 made that old content
 *       visible again as a duplicate column and stray percent-formatted
 *       value, without this code ever having written it. refreshDashboard
 *       Summary_ now does a full clearContent()+clearFormat() over the whole
 *       region immediately after breakApart(), so anything revealed gets
 *       wiped before the section is rebuilt from scratch every time.
 *
 * v6.5: BUGFIX - the v6.4 breakApart() fix was itself too narrow. Google
 *       Sheets applies the same "must select all cells in a merged range"
 *       restriction to breakApart() as to merge() - if an existing merge is
 *       wider than the range you call breakApart() on, it throws the same
 *       error instead of clearing it. Widened every breakApart() call to a
 *       full row (columns A-Z) so it's guaranteed to fully contain whatever
 *       merge shape actually exists. Also added the same defensive
 *       breakApart() to setupOneDormSheet_ and markInactiveDormSheet_ -
 *       setupOneDormSheet_ in particular runs on every active dorm on every
 *       registry-changing action, making it the most exposed spot in the
 *       whole codebase to this class of bug.
 *
 * v6.4: BUGFIX - refreshDashboardSummary_'s merge() call could throw "You
 *       must select all cells in a merged range to merge or unmerge them"
 *       if the Dashboard's A15:H15 (or anything in the row 20+ table area)
 *       was merged in a shape that didn't exactly match, which aborted the
 *       whole calling function (Add Dorm, Deactivate Dorm, Setup) partway
 *       through - explaining "nothing else in the dashboard is added".
 *       Both merge-affecting regions now call breakApart() first, which is
 *       a no-op on ordinary unmerged cells, so this can't happen again.
 *
 * v6.3: Add Dorm / Deactivate Dorm now keep the Dashboard's "Current Summary"
 *       per-dorm table in sync (new refreshDashboardSummary_ - also runs on
 *       Setup). That table predates this whole project and nothing used to
 *       touch it, which is why a newly added dorm never got a row. Also
 *       replaced the "Dorms Tracked" figure, previously a typed-in static
 *       number, with a live formula against the Settings dorm table.
 *
 * v6.2: BUGFIX - Events Attended and Eligible Events could show 1 for a
 *       student with zero real attendance/eligible events. Cause: Google
 *       Sheets' COUNTUNIQUE() counts a #N/A error as "1 unique value"
 *       instead of propagating it, so IFERROR(COUNTUNIQUE(FILTER(...)),0)
 *       never caught the "FILTER found nothing" case the way it was assumed
 *       to. Rebuilt Events Attended / Eligible Events with SUMPRODUCT over
 *       bounded ranges instead of FILTER/COUNTUNIQUE, which returns a
 *       genuine 0 for no matches rather than an error to swallow. Last
 *       Attended was unaffected (MAX does propagate errors correctly) and is
 *       unchanged.
 *
 * v6: Eligible Events and Events Attended are now dorm-aware. Each event
 *     sheet gets an "Applies To Dorm(s)" field, auto-detected from the event
 *     name (e.g. "Dorm D CCM" -> Dorm D; "Dorm A/C/K CCM & Movie Night" ->
 *     Dorm A, Dorm C, Dorm K; a name with no dorm mentioned -> every dorm),
 *     editable any time. EventLog and AttendanceLog carry this scope through
 *     to the formulas, which now only count an event toward a student's
 *     Eligible/Attended totals if it's blank (applies to everyone) or lists
 *     their dorm - applied symmetrically to both numbers so an out-of-scope
 *     event can never push Attendance % over 100%. Data Audit flags any
 *     scope that references a dorm that isn't actually configured yet.
 *     Existing EventLog/AttendanceLog sheets are migrated in place (a new
 *     column is added, nothing is cleared) and existing event sheets get
 *     their scope backfilled automatically the next time Setup runs.
 *
 * v5.2: New Delete Student (Permanent) - distinct from Deactivate. For
 *       mistaken/duplicate registry entries: removes the registry row
 *       outright, archives their AttendanceLog rows instead of destroying
 *       them, and logs what was deleted (and when) to a new DeletedStudents
 *       sheet. Requires typing the student ID twice to confirm.
 *
 * v5.1: Imported events now sync into AttendanceLog immediately at creation
 *       time (previously only the auto-refresh trigger or a manual Refresh
 *       Attendance Log picked them up, which could lag). Data Audit now
 *       flags attendance recorded before a student's own registered Start
 *       Date, since that silently zeroes out their Events Attended count.
 *
 * v5: Reusable copy-paste templates for every bulk action, per-student
 *     attendance stats on StudentRegistry, Event IDs that combine the date
 *     with a 2-digit sequence, a trimmed Maintenance menu, and an auto-
 *     refresh trigger so Eligible Events / Attendance % update as soon as
 *     attendee IDs are typed or pasted (not just when Refresh Attendance
 *     Log is run manually).
 *
 * Install:
 * 1. Open the imported Google Sheet.
 * 2. Extensions -> Apps Script.
 * 3. Paste this file into Code.gs (replace everything).
 * 4. Save.
 * 5. Reload the Sheet.
 * 6. Run Dorm Attendance -> Maintenance -> Setup / Repair Workbook once.
 *
 * Upgrading from v6.9 to v7.0 (one time only):
 * a. Paste this file into Code.gs (replace everything).
 * b. In the Apps Script editor click + -> HTML, name it exactly
 *    ImportInviteOnlyEvent (no .html), and paste that file's contents in.
 * c. Save, reload the Sheet, run Maintenance -> Setup / Repair Workbook once.
 *    That adds the new log columns, updates every Eligible Events formula,
 *    and appends section 4 to BulkActionTemplates without touching anything
 *    already typed there.
 *
 * Upgrading from v6.8 or earlier (one time only):
 * a. In the Apps Script editor, open Triggers (clock icon on the left) and
 *    delete the old "handleAttendanceEdit" and "handleSpreadsheetChange"
 *    triggers. They are replaced by the built-in onEdit() below and would
 *    otherwise make the refresh run twice (or fail with errors).
 * b. Optional but recommended: remove the old, broader grant at
 *    https://myaccount.google.com/permissions ("Attendance Tracker Project")
 *    so the next run asks only for the reduced permissions.
 */

const DEFAULT_DORMS = ['Dorm A', 'Dorm B', 'Dorm D', 'Dorm E', 'Meikokan'];
const DORM_TABLE_START_ROW = 10;
const DORM_TABLE_HEADERS = ['Dorm Name', 'Status', 'Created At', 'Deactivated At', 'Notes'];
const DASHBOARD_SHEET = 'Dashboard';
const REGISTRY_SHEET = 'StudentRegistry';
const EVENT_LOG_SHEET = 'EventLog';
const ATTENDANCE_LOG_SHEET = 'AttendanceLog';
const SETTINGS_SHEET = 'Settings';
const TEMPLATE_SHEET = 'EventTemplate';
const TRANSFER_LOG_SHEET = 'TransferLog';
const DELETED_STUDENTS_SHEET = 'DeletedStudents';
const EVENT_LOG_ARCHIVE_SHEET = 'EventLog_Archive';
const ATTENDANCE_LOG_ARCHIVE_SHEET = 'AttendanceLog_Archive';
const DATA_AUDIT_SHEET = 'DataAudit';
const BULK_TEMPLATES_SHEET = 'BulkActionTemplates';
const EVENT_PREFIX = 'EVT_';
const ATTENDEE_START_ROW = 8;
const ATTENDEE_COL = 1; // Column A
const MAX_EVENT_ID_NUMBER = 99; // The sequence part of the Event ID is kept to 2 digits (...-01 to ...-99).

// v7.0: invite-only events.
const INVITE_ONLY_TYPE = 'Invite-Only';     // Value of B6 on an invite-only event tab, and of the Event Type log columns.
const INVITE_ATTENDED_COL = 2;              // Column B - the "Attended?" checkbox on invite-only event tabs.
const INVITE_CHECKBOX_BUFFER_ROWS = 30;     // Spare checkbox rows below the invite list for late additions.
const INVITE_TEMPLATE_TITLE = '4. Import Invite-Only Event  \u2014  Events -> Import Invite-Only Event (Tick Attendance)';

// v7.0: one place for the log layouts (EventLog gained G, AttendanceLog gained J).
const EVENT_LOG_HEADERS = ['Event ID', 'Event Name', 'Event Date', 'Sheet Name', 'Created At', 'Applies To', 'Event Type'];
const ATTENDANCE_LOG_HEADERS = ['Record ID', 'Event ID', 'Event Date', 'Event Name', 'Student ID', 'Dorm', 'Event Sheet', 'Logged At', 'Applies To', 'Event Type'];
const EVENT_LOG_COLS = EVENT_LOG_HEADERS.length;           // 7
const ATTENDANCE_LOG_COLS = ATTENDANCE_LOG_HEADERS.length; // 10

function onOpen() {
  const ui = SpreadsheetApp.getUi();

  const studentsMenu = ui.createMenu('Students')
    .addItem('Register Student (Single)', 'registerStudent')
    .addItem('Bulk Register (Paste from Excel)', 'openBulkRegisterDialog')
    .addItem('Transfer Student (Change Dorm)', 'transferStudentDorm')
    .addSeparator()
    .addItem('Deactivate Student', 'deactivateStudent')
    .addItem('Bulk Deactivate (Paste List)', 'openBulkDeactivateDialog')
    .addSeparator()
    .addItem('Delete Student (Permanent)', 'deleteStudent');

  const eventsMenu = ui.createMenu('Events')
    .addItem('Create New Event Sheet', 'createNewEventSheet')
    .addItem('Import Event with Attendance Details', 'openImportEventDialog')
    .addItem('Import Invite-Only Event (Tick Attendance)', 'openImportInviteOnlyDialog')
    .addSeparator()
    .addItem('Refresh Attendance Log', 'refreshAttendanceLog')
    .addItem('Repair Event Sheet Formulas', 'repairEventSheetFormulas')
    .addSeparator()
    .addItem('Archive Previous Years\u2019 Events', 'archivePreviousYearEvents')
    .addItem('Clean Up Deleted Event References', 'cleanUpDeletedEventReferences');

  const dormsMenu = ui.createMenu('Dorms')
    .addItem('Add Dorm', 'addDorm')
    .addItem('Deactivate Dorm', 'deactivateDorm')
    .addSeparator()
    .addItem('Rebuild Dorm Summary Sheets', 'rebuildDormSummarySheets');

  const maintenanceMenu = ui.createMenu('Maintenance')
    .addItem('Run Data Audit', 'runDataAudit')
    .addItem('Protect System Sheets', 'protectSystemSheets')
    .addSeparator()
    .addItem('Reset Bulk Action Templates', 'resetBulkActionTemplates')
    .addSeparator()
    .addItem('Setup / Repair Workbook', 'setupWorkbook');

  ui.createMenu('Dorm Attendance')
    .addSubMenu(studentsMenu)
    .addSubMenu(eventsMenu)
    .addSubMenu(dormsMenu)
    .addSubMenu(maintenanceMenu)
    .addToUi();

  // v6.9: replaces the old installable onChange trigger. If an event tab was
  // deleted since the last visit, drop its leftover EventLog/AttendanceLog
  // rows now. Cheap when there's nothing to remove. Wrapped in try/catch so
  // a view-only user (who can't write) still gets the menu without errors.
  try {
    pruneOrphanedEventLogEntries_();
  } catch (err) {
    // Ignore - the menu is already in place, and the next refresh cleans up anyway.
  }
}

function setupWorkbook() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  ensureSheet_(DASHBOARD_SHEET);
  ensureSheet_(REGISTRY_SHEET);
  ensureSheet_(EVENT_LOG_SHEET);
  ensureSheet_(ATTENDANCE_LOG_SHEET);
  ensureSheet_(SETTINGS_SHEET);
  ensureSheet_(TEMPLATE_SHEET);
  ensureSheet_(BULK_TEMPLATES_SHEET);
  setupSettings_();
  const activeDorms = getActiveDormNames_();
  activeDorms.forEach(name => ensureSheet_(name));
  setupRegistry_();
  setupEventLog_();
  setupAttendanceLog_();
  setupEventTemplate_();
  setupTransferLog_();
  setupBulkActionTemplates_();
  setupDormSheets_();

  // Backfill event-sheet metadata (Venue/Expected/Reported/Applies To) on
  // the template and every existing event sheet, so upgrading to a newer
  // version only takes "paste + run Setup" - no separate repair step needed.
  const template = ss.getSheetByName(TEMPLATE_SHEET);
  if (template) ensureEventSheetMetaLabels_(template);
  ss.getSheets().forEach(sheet => {
    if (isEventSheet_(sheet)) ensureEventSheetMetaLabels_(sheet);
  });

  rebuildDormSummarySheets();
  refreshDashboardInstructions();
  refreshDashboardSummary_();
  // v6.9: no installable triggers any more - auto-refresh uses the simple
  // onEdit() trigger below, which needs no extra permission to set up.
  pruneOrphanedEventLogEntries_();

  // Hide backend/template/log sheets. Editors can still unhide them from the
  // sheet tab bar; run Maintenance -> Protect System Sheets for real access control.
  // BulkActionTemplates is deliberately NOT hidden - it's a working sheet staff use directly.
  [TEMPLATE_SHEET, EVENT_LOG_SHEET, ATTENDANCE_LOG_SHEET, SETTINGS_SHEET, TRANSFER_LOG_SHEET].forEach(name => {
    const sh = ss.getSheetByName(name);
    if (sh && !sh.isSheetHidden()) sh.hideSheet();
  });
  [EVENT_LOG_ARCHIVE_SHEET, ATTENDANCE_LOG_ARCHIVE_SHEET, DATA_AUDIT_SHEET, DELETED_STUDENTS_SHEET].forEach(name => {
    const sh = ss.getSheetByName(name);
    if (sh && !sh.isSheetHidden()) sh.hideSheet();
  });

  SpreadsheetApp.getUi().alert(
    'Setup complete. The Dorm Attendance menu is ready.\n\n' +
    'Once your staff list is final, run Maintenance -> Protect System Sheets to lock backend sheets and formula columns.'
  );
}

function createNewEventSheet() {
  const ui = SpreadsheetApp.getUi();
  const ss = SpreadsheetApp.getActiveSpreadsheet();

  const nameResponse = ui.prompt(
    'Create New Event Sheet',
    'Enter the event name:',
    ui.ButtonSet.OK_CANCEL
  );
  if (nameResponse.getSelectedButton() !== ui.Button.OK) return;
  const eventName = nameResponse.getResponseText().trim();
  if (!eventName) {
    ui.alert('Event name is required.');
    return;
  }

  const dateResponse = ui.prompt(
    'Create New Event Sheet',
    'Enter the event date in yyyy-mm-dd format. The generated sheet will also have a date-picker cell in B4.',
    ui.ButtonSet.OK_CANCEL
  );
  if (dateResponse.getSelectedButton() !== ui.Button.OK) return;
  const eventDate = parseDate_(dateResponse.getResponseText().trim());
  if (!eventDate) {
    ui.alert('Invalid date. Please use yyyy-mm-dd format, for example 2026-07-11.');
    return;
  }

  let eventId;
  try {
    eventId = makeEventId_(eventDate);
  } catch (err) {
    ui.alert(err.message);
    return;
  }
  const sheetNameBase = EVENT_PREFIX + Utilities.formatDate(eventDate, Session.getScriptTimeZone(), 'yyyyMMdd') + '_' + sanitizeSheetName_(eventName);
  const sheetName = uniqueSheetName_(sheetNameBase.substring(0, 90));

  const detectedScope = detectDormScope_(eventName);
  const scopeResponse = ui.prompt(
    'Create New Event Sheet',
    'Which dorm(s) does this event apply to? Detected from the name: ' + (detectedScope ? detectedScope.split(',').join(', ') : 'All dorms') +
    '.\n\nLeave blank to accept that, type "All" for every dorm, or type dorm names separated by commas to override.',
    ui.ButtonSet.OK_CANCEL
  );
  if (scopeResponse.getSelectedButton() !== ui.Button.OK) return;
  const scopeText = scopeResponse.getResponseText().trim();
  const dormScope = scopeText ? normalizeDormScope_(scopeText) : detectedScope;

  let template = ss.getSheetByName(TEMPLATE_SHEET);
  if (!template) {
    setupEventTemplate_();
    template = ss.getSheetByName(TEMPLATE_SHEET);
  }

  const newSheet = template.copyTo(ss).setName(sheetName);
  ss.setActiveSheet(newSheet);
  ss.moveActiveSheet(ss.getNumSheets());

  newSheet.getRange('B3').setValue(eventId);
  newSheet.getRange('B4').setValue(eventDate).setNumberFormat('yyyy-mm-dd');
  newSheet.getRange('B5').setValue(eventName);
  newSheet.getRange('E6').setValue(dormScope);
  applyEventSheetFormatting_(newSheet);

  upsertEventLog_(eventId, eventName, eventDate, sheetName, dormScope);

  ui.alert('Created event sheet: ' + sheetName + ' (Event ID ' + eventId + ')' +
    '\nApplies to: ' + (dormScope ? dormScope.split(',').join(', ') : 'All dorms'));
}

function registerStudent() {
  const ui = SpreadsheetApp.getUi();

  const idResponse = ui.prompt('Register Student', 'Enter the student ID number:', ui.ButtonSet.OK_CANCEL);
  if (idResponse.getSelectedButton() !== ui.Button.OK) return;
  const studentId = normalizeId_(idResponse.getResponseText());
  if (!studentId) {
    ui.alert('Student ID is required.');
    return;
  }

  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sh = ss.getSheetByName(REGISTRY_SHEET);
  const row = findRegistryRow_(studentId);
  const existingDorm = row ? String(sh.getRange(row, 3).getValue()).trim() : '';
  const existingStartDate = row ? sh.getRange(row, 5).getValue() : null;

  const dormResponse = ui.prompt(
    'Register Student',
    'Enter dorm exactly as one of the following active dorms:\n' + getActiveDormNames_().join(', '),
    ui.ButtonSet.OK_CANCEL
  );
  if (dormResponse.getSelectedButton() !== ui.Button.OK) return;
  const dorm = dormResponse.getResponseText().trim();
  const activeDorms = getActiveDormNames_();
  if (activeDorms.indexOf(dorm) === -1) {
    ui.alert('Invalid dorm. Please use one of the active dorms: ' + activeDorms.join(', '));
    return;
  }

  if (row && existingDorm && existingDorm !== dorm) {
    const proceed = ui.alert(
      'Register Student',
      studentId + ' is already registered in ' + existingDorm + '. Registering here as a fresh entry will reset ' +
      'their attendance start date unless you type one in the next step.\n\n' +
      'Tip: use Students -> Transfer Student instead to move dorms while keeping their attendance history.\n\n' +
      'Continue anyway?',
      ui.ButtonSet.YES_NO
    );
    if (proceed !== ui.Button.YES) return;
  }

  const notesResponse = ui.prompt(
    'Register Student',
    'Optional: enter student name or notes. Leave blank if student IDs should remain the only identifiers.',
    ui.ButtonSet.OK_CANCEL
  );
  if (notesResponse.getSelectedButton() !== ui.Button.OK) return;
  const notes = notesResponse.getResponseText().trim();

  const startResponse = ui.prompt(
    'Register Student',
    'Enter start date in yyyy-mm-dd format, or leave blank to keep the existing start date (new students default to today):',
    ui.ButtonSet.OK_CANCEL
  );
  if (startResponse.getSelectedButton() !== ui.Button.OK) return;
  const startText = startResponse.getResponseText().trim();

  let startDate;
  if (startText) {
    startDate = parseDate_(startText);
    if (!startDate) {
      ui.alert('Invalid start date. Please use yyyy-mm-dd format.');
      return;
    }
  } else if (row && existingStartDate instanceof Date) {
    startDate = existingStartDate; // Preserve history instead of silently resetting it.
  } else {
    startDate = new Date();
  }

  const now = new Date();

  if (row) {
    sh.getRange(row, 2, 1, 5).setValues([[notes, dorm, 'Active', startDate, now]]);
  } else {
    sh.appendRow([studentId, notes, dorm, 'Active', startDate, now]);
  }

  rebuildDormSummarySheets();
  ui.alert('Registered student ' + studentId + ' in ' + dorm + '.');
}

function deactivateStudent() {
  const ui = SpreadsheetApp.getUi();
  const idResponse = ui.prompt('Deactivate Student', 'Enter the student ID number:', ui.ButtonSet.OK_CANCEL);
  if (idResponse.getSelectedButton() !== ui.Button.OK) return;
  const studentId = normalizeId_(idResponse.getResponseText());
  if (!studentId) {
    ui.alert('Student ID is required.');
    return;
  }

  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sh = ss.getSheetByName(REGISTRY_SHEET);
  const row = findRegistryRow_(studentId);
  if (!row) {
    ui.alert('Student ID was not found in the registry.');
    return;
  }

  sh.getRange(row, 4).setValue('Inactive');
  sh.getRange(row, 6).setValue(new Date());
  rebuildDormSummarySheets();
  ui.alert('Deactivated student ' + studentId + '. Historical attendance records were preserved.');
}

/**
 * Permanently removes a student from StudentRegistry - for mistaken or
 * duplicate entries, not for students who actually left. Distinct from
 * Deactivate Student, which keeps everything and only flips Status.
 *
 * What happens:
 * - The registry row is deleted outright (deleteRow), not just marked Inactive.
 * - Their AttendanceLog rows are moved to AttendanceLog_Archive first (kept
 *   for reference, but no longer counted toward anyone's stats) rather than
 *   being destroyed outright.
 * - A copy of the deleted row is logged to DeletedStudents (who/what/when),
 *   so there's a record of exactly what was removed.
 * - TransferLog entries are left alone - they're an audit trail independent
 *   of whether the registry row still exists.
 * - Event sheets are NOT edited automatically: if the ID was typed as an
 *   attendee somewhere, that row is left as-is (the confirmation dialog
 *   lists which sheets, if any, so staff can clean those up by hand; Data
 *   Audit will also flag it as "Attendee not found in registry" afterward).
 */
function deleteStudent() {
  const ui = SpreadsheetApp.getUi();
  const idResponse = ui.prompt('Delete Student (Permanent)', 'Enter the student ID to permanently delete:', ui.ButtonSet.OK_CANCEL);
  if (idResponse.getSelectedButton() !== ui.Button.OK) return;
  const studentId = normalizeId_(idResponse.getResponseText());
  if (!studentId) {
    ui.alert('Student ID is required.');
    return;
  }

  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sh = ss.getSheetByName(REGISTRY_SHEET);
  const row = findRegistryRow_(studentId);
  if (!row) {
    ui.alert('Student ID was not found in the registry.');
    return;
  }

  const rowValues = sh.getRange(row, 1, 1, 6).getValues()[0]; // Student ID, Notes, Dorm, Status, Start Date, Last Updated
  const dorm = String(rowValues[2]).trim();
  const status = String(rowValues[3]).trim();
  const attCount = countAttendanceRowsForStudent_(ATTENDANCE_LOG_SHEET, studentId);
  const eventSheetsWithId = findEventSheetsContainingAttendee_(studentId);

  let warning = 'Permanently delete ' + studentId + ' (' + dorm + ', ' + status + ')?\n\n' +
    '\u2022 Removed from StudentRegistry entirely - this cannot be undone from the menu.\n' +
    '\u2022 ' + attCount + ' attendance record(s) will be moved to AttendanceLog_Archive (kept, but no longer counted for anyone).\n' +
    '\u2022 A copy of their registry row is saved to the DeletedStudents log first.\n' +
    '\u2022 Dorm transfer history, if any, is left as-is.';
  if (eventSheetsWithId.length) {
    warning += '\n\nNote: this ID still appears as a typed attendee on: ' + eventSheetsWithId.join(', ') +
      '. Those rows are NOT removed automatically - edit those sheets by hand if you want them gone too.';
  }
  warning += '\n\nThis is meant for mistaken or duplicate entries. If this student actually left the dorm, use Deactivate Student instead to keep their history.';

  const confirm = ui.alert('Delete Student (Permanent)', warning, ui.ButtonSet.YES_NO);
  if (confirm !== ui.Button.YES) return;

  const confirmId = ui.prompt('Delete Student (Permanent)', 'Type the student ID again to confirm permanent deletion:', ui.ButtonSet.OK_CANCEL);
  if (confirmId.getSelectedButton() !== ui.Button.OK || normalizeId_(confirmId.getResponseText()) !== studentId) {
    ui.alert('Student ID did not match. Nothing was deleted.');
    return;
  }

  logDeletedStudent_(rowValues, attCount);
  archiveAttendanceForStudent_(studentId);
  sh.deleteRow(row);

  rebuildDormSummarySheets();
  ui.alert('Deleted ' + studentId + '. ' + attCount + ' attendance record(s) moved to AttendanceLog_Archive.');
}

function countAttendanceRowsForStudent_(sheetName, studentId) {
  const sh = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(sheetName);
  if (!sh || sh.getLastRow() < 2) return 0;
  const ids = sh.getRange(2, 5, sh.getLastRow() - 1, 1).getValues().flat(); // column E = Student ID
  return ids.filter(v => normalizeId_(v) === studentId).length;
}

function findEventSheetsContainingAttendee_(studentId) {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const found = [];
  ss.getSheets().forEach(sheet => {
    if (!isEventSheet_(sheet)) return;
    const lastRow = Math.max(sheet.getLastRow(), ATTENDEE_START_ROW);
    const numRows = lastRow - ATTENDEE_START_ROW + 1;
    const ids = sheet.getRange(ATTENDEE_START_ROW, ATTENDEE_COL, numRows, 1).getValues().flat();
    if (ids.some(v => normalizeId_(v) === studentId)) found.push(sheet.getName());
  });
  return found;
}

function logDeletedStudent_(rowValues, attendanceCount) {
  const sh = ensureSheet_(DELETED_STUDENTS_SHEET);
  if (sh.getLastRow() === 0 || sh.getRange('A1').getValue() !== 'Student ID') {
    sh.clear();
    sh.getRange('A1:G1').setValues([['Student ID', 'Notes', 'Dorm', 'Status', 'Start Date', 'Attendance Records Archived', 'Deleted At']]);
    styleHeader_(sh.getRange('A1:G1'));
    sh.getRange('E2:E').setNumberFormat('yyyy-mm-dd hh:mm');
    sh.getRange('G2:G').setNumberFormat('yyyy-mm-dd hh:mm');
    sh.setFrozenRows(1);
  }
  sh.appendRow([rowValues[0], rowValues[1], rowValues[2], rowValues[3], rowValues[4], attendanceCount, new Date()]);
}

/** Moves every AttendanceLog row for one student into AttendanceLog_Archive, removing them from the live log. */
function archiveAttendanceForStudent_(studentId) {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const attSh = ss.getSheetByName(ATTENDANCE_LOG_SHEET);
  if (!attSh || attSh.getLastRow() < 2) return;

  const data = attSh.getRange(2, 1, attSh.getLastRow() - 1, ATTENDANCE_LOG_COLS).getValues();
  const keep = [];
  const toArchive = [];
  data.forEach(r => {
    if (normalizeId_(r[4]) === studentId) toArchive.push(r);
    else keep.push(r);
  });
  if (!toArchive.length) return;

  const attArchive = ensureSheet_(ATTENDANCE_LOG_ARCHIVE_SHEET);
  ensureAttendanceArchiveHeader_(attArchive);

  attArchive.getRange(attArchive.getLastRow() + 1, 1, toArchive.length, ATTENDANCE_LOG_COLS).setValues(toArchive);
  attSh.getRange(2, 1, attSh.getMaxRows() - 1, ATTENDANCE_LOG_COLS).clearContent();
  if (keep.length) attSh.getRange(2, 1, keep.length, ATTENDANCE_LOG_COLS).setValues(keep);
}

/**
 * v7.0: shared header setup for the two archive sheets. Writes the full,
 * current header row every time (harmless - existing column names are
 * unchanged, only new ones get added on the right), so archives created by
 * older versions pick up "Applies To" / "Event Type" automatically.
 */
function ensureAttendanceArchiveHeader_(attArchive) {
  const isNew = attArchive.getLastRow() < 1;
  attArchive.getRange(1, 1, 1, ATTENDANCE_LOG_COLS).setValues([ATTENDANCE_LOG_HEADERS]);
  styleHeader_(attArchive.getRange(1, 1, 1, ATTENDANCE_LOG_COLS));
  if (isNew) {
    attArchive.getRange('C:C').setNumberFormat('yyyy-mm-dd');
    attArchive.getRange('H:H').setNumberFormat('yyyy-mm-dd hh:mm');
    attArchive.setFrozenRows(1);
  }
}

function ensureEventArchiveHeader_(evArchive) {
  const isNew = evArchive.getLastRow() < 1;
  evArchive.getRange(1, 1, 1, EVENT_LOG_COLS).setValues([EVENT_LOG_HEADERS]);
  styleHeader_(evArchive.getRange(1, 1, 1, EVENT_LOG_COLS));
  if (isNew) {
    evArchive.getRange('C:C').setNumberFormat('yyyy-mm-dd');
    evArchive.getRange('E:E').setNumberFormat('yyyy-mm-dd hh:mm');
    evArchive.setFrozenRows(1);
  }
}

function addDorm() {
  const ui = SpreadsheetApp.getUi();
  const response = ui.prompt('Add Dorm', 'Enter the new dorm name exactly as it should appear on its summary tab:', ui.ButtonSet.OK_CANCEL);
  if (response.getSelectedButton() !== ui.Button.OK) return;
  const dormName = response.getResponseText().trim();
  if (!dormName) {
    ui.alert('Dorm name is required.');
    return;
  }
  if (isReservedSheetName_(dormName)) {
    ui.alert('That name is reserved for a system sheet. Please use a different dorm name.');
    return;
  }

  const notesResponse = ui.prompt('Add Dorm', 'Optional: enter notes for this dorm. Leave blank if unnecessary.', ui.ButtonSet.OK_CANCEL);
  if (notesResponse.getSelectedButton() !== ui.Button.OK) return;
  const notes = notesResponse.getResponseText().trim();

  const result = upsertDormSetting_(dormName, 'Active', notes);
  const sh = ensureSheet_(dormName);
  setupOneDormSheet_(sh, dormName);
  setupRegistry_();
  rebuildDormSummarySheets();
  refreshDashboardSummary_();
  ui.alert((result === 'reactivated' ? 'Reactivated dorm: ' : 'Added dorm: ') + dormName + '.');
}

function deactivateDorm() {
  const ui = SpreadsheetApp.getUi();
  const activeDorms = getActiveDormNames_();
  if (!activeDorms.length) {
    ui.alert('There are no active dorms to deactivate.');
    return;
  }
  const response = ui.prompt('Deactivate Dorm', 'Enter the dorm name to deactivate. Active dorms:\n' + activeDorms.join(', '), ui.ButtonSet.OK_CANCEL);
  if (response.getSelectedButton() !== ui.Button.OK) return;
  const dormName = response.getResponseText().trim();
  if (activeDorms.indexOf(dormName) === -1) {
    ui.alert('Dorm not found among active dorms. No changes were made.');
    return;
  }

  const confirm = ui.alert(
    'Deactivate Dorm',
    'Deactivate "' + dormName + '"?\n\nThis will remove it from active dorm dropdowns and future summary rebuilding. The dorm sheet, student registry rows, EventLog, and AttendanceLog will be preserved.',
    ui.ButtonSet.YES_NO
  );
  if (confirm !== ui.Button.YES) return;

  markDormInactive_(dormName);
  markInactiveDormSheet_(dormName);
  setupRegistry_();
  rebuildDormSummarySheets();
  refreshDashboardSummary_();
  ui.alert('Deactivated dorm: ' + dormName + '. Historical records were preserved.');
}

function repairEventSheetFormulas() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  let repaired = 0;

  // Repair the reusable template first.
  const template = ss.getSheetByName(TEMPLATE_SHEET);
  if (template) {
    applyEventLookupFormulas_(template);
    ensureEventSheetMetaLabels_(template);
    repaired++;
  }

  // Repair all existing event sheets, including ones created before this update.
  ss.getSheets().forEach(sheet => {
    if (!isEventSheet_(sheet)) return;
    applyEventLookupFormulas_(sheet);
    ensureEventSheetMetaLabels_(sheet);
    repaired++;
  });

  SpreadsheetApp.getUi().alert('Event lookup formulas repaired on ' + repaired + ' sheet(s).');
}

/** Backfills the event-sheet metadata fields for sheets created before v4/v6. Leaves any existing values alone - only adds what's missing. */
function ensureEventSheetMetaLabels_(sh) {
  if (String(sh.getRange('D3').getValue()).trim() !== 'Venue') {
    const existingVenue = sh.getRange('E3').getValue();
    const existingExpected = sh.getRange('E4').getValue();
    const existingReported = sh.getRange('E5').getValue();
    sh.getRange('D3:E5').setValues([
      ['Venue', existingVenue || ''],
      ['Expected Attendees', existingExpected || ''],
      ['Reported Attendee Count', existingReported || '']
    ]);
    sh.getRange('D3:D5').setBackground('#D9EAF7').setFontWeight('bold');
    sh.getRange('E3:E5').setBackground('#FFF2CC');
  }

  if (String(sh.getRange('D6').getValue()).trim() !== 'Applies To Dorm(s)') {
    const detected = detectDormScope_(sh.getRange('B5').getValue());
    sh.getRange('D6:E6').setValues([['Applies To Dorm(s)', detected]]);
    sh.getRange('D6').setBackground('#D9EAF7').setFontWeight('bold');
    sh.getRange('E6').setBackground('#FFF2CC');
    sh.getRange('E6').setNote('Leave blank (or type "All") for an event open to every dorm. Otherwise, comma-separate the dorm names it applies to, e.g. "Dorm A, Dorm C". Auto-filled from the event name when possible - edit any time.');
  }
}

function refreshAttendanceLog() {
  const result = refreshAttendanceLogSilently_();
  const message = 'Attendance log refreshed.\n\nEvents found: ' + result.eventCount +
    '\nAttendance records: ' + result.attendanceCount +
    (result.warnings.length ? '\n\nWarnings:\n' + result.warnings.slice(0, 20).join('\n') + (result.warnings.length > 20 ? '\n...and more.' : '') : '');
  SpreadsheetApp.getUi().alert(message);
}

/**
 * Core of Refresh Attendance Log, without the UI alert, so it can also run
 * silently from the auto-refresh on edit (see onEdit below).
 *
 * Note: this rebuilds EventLog and AttendanceLog entirely from the event
 * tabs that currently exist, so any tab that was deleted drops out of both
 * logs automatically on the next refresh.
 */
function refreshAttendanceLogSilently_() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const registry = getRegistryMap_();
  const attendanceRows = [];
  const eventRows = [];
  const now = new Date();
  const warnings = [];

  ss.getSheets().forEach(sheet => {
    if (!isEventSheet_(sheet)) return;

    const eventId = String(sheet.getRange('B3').getValue()).trim();
    const eventDate = sheet.getRange('B4').getValue();
    const eventName = String(sheet.getRange('B5').getValue()).trim();
    const sheetName = sheet.getName();

    // v7.0: invite-only tabs ignore the dorm scope (the invite list IS the
    // scope) and only log residents whose "Attended?" box is ticked.
    const inviteOnly = isInviteOnlySheet_(sheet);
    const dormScope = inviteOnly ? '' : normalizeDormScope_(sheet.getRange('E6').getValue());
    const eventType = inviteOnly ? INVITE_ONLY_TYPE : '';

    if (!eventId || !(eventDate instanceof Date) || !eventName) {
      warnings.push(sheetName + ': missing Event ID, Event Date, or Event Name.');
      return;
    }

    eventRows.push([eventId, eventName, eventDate, sheetName, now, dormScope, eventType]);

    const lastRow = Math.max(sheet.getLastRow(), ATTENDEE_START_ROW);
    const numRows = lastRow - ATTENDEE_START_ROW + 1;
    const numCols = inviteOnly ? INVITE_ATTENDED_COL - ATTENDEE_COL + 1 : 1;
    const block = sheet.getRange(ATTENDEE_START_ROW, ATTENDEE_COL, numRows, numCols).getValues();
    const seenInEvent = {};

    block.forEach(r => {
      const studentId = normalizeId_(r[0]);
      if (!studentId) return;
      // Unticked invitees are skipped BEFORE being marked as seen, so if the
      // same ID is listed twice and only one copy is ticked, it still counts.
      if (inviteOnly && !isTicked_(r[INVITE_ATTENDED_COL - ATTENDEE_COL])) return;
      if (seenInEvent[studentId]) return;
      seenInEvent[studentId] = true;

      const reg = registry[studentId];
      if (!reg) {
        warnings.push(sheetName + ': unregistered ID skipped: ' + studentId);
        return;
      }

      const recordId = eventId + '|' + studentId;
      attendanceRows.push([recordId, eventId, eventDate, eventName, studentId, reg.dorm, sheetName, now, dormScope, eventType]);
    });
  });

  writeEventLog_(eventRows);
  writeAttendanceLog_(attendanceRows);
  rebuildDormSummarySheets();
  setLastRefreshed_(now);

  return { eventCount: eventRows.length, attendanceCount: attendanceRows.length, warnings: warnings };
}

/**
 * Simple onEdit trigger (v6.9 - replaces the old installable
 * handleAttendanceEdit trigger, which needed the "run when you are not
 * present" permission). Google runs a function named onEdit automatically
 * whenever someone edits the sheet, with no extra permission and nothing to
 * install.
 *
 * Whenever someone types or pastes into the Student ID column of an event
 * sheet, this silently re-runs the attendance refresh so Events Attended /
 * Eligible Events / Attendance % (on the dorm sheets AND on StudentRegistry)
 * update right away instead of waiting for someone to remember to click
 * Events -> Refresh Attendance Log.
 *
 * Limits of simple triggers to be aware of: each run gets 30 seconds (the
 * old installable trigger got 6 minutes), and they don't fire for edits made
 * by the script itself (the import path already refreshes explicitly). If
 * the workbook ever grows large enough that the auto-refresh stops keeping
 * up, Events -> Refresh Attendance Log still works as before.
 */
function onEdit(e) {
  if (!e || !e.range) return;
  const sheet = e.range.getSheet();
  if (e.range.getColumn() > INVITE_ATTENDED_COL) return; // Only A (IDs) or B (invite-only ticks) can matter.
  if (e.range.getRow() + e.range.getNumRows() - 1 < ATTENDEE_START_ROW) return; // Ignore edits above the attendee list.
  if (!isEventSheet_(sheet)) return;

  // v7.0: on invite-only tabs, ticking/unticking the "Attended?" box in
  // column B also triggers a refresh. On regular tabs only column A matters.
  const inviteOnly = isInviteOnlySheet_(sheet);
  if (!inviteOnly && e.range.getColumn() > ATTENDEE_COL) return;

  // If someone types/pastes a late invitee into column A, make sure that row
  // has an "Attended?" checkbox too (adds the checkbox rule only - never
  // changes an existing tick).
  if (inviteOnly && e.range.getColumn() === ATTENDEE_COL) {
    const firstRow = Math.max(e.range.getRow(), ATTENDEE_START_ROW);
    const lastRow = e.range.getRow() + e.range.getNumRows() - 1;
    applyInviteCheckboxes_(sheet, firstRow, lastRow - firstRow + 1);
  }

  refreshAttendanceLogSilently_();
}

function rebuildDormSummarySheets() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const registry = getRegistryRows_();

  const activeDorms = getActiveDormNames_();
  activeDorms.forEach(dorm => {
    const sh = ss.getSheetByName(dorm) || ss.insertSheet(dorm);
    setupOneDormSheet_(sh, dorm);

    const rows = registry
      .filter(r => r.status === 'Active' && r.dorm === dorm)
      .sort((a, b) => String(a.studentId).localeCompare(String(b.studentId)))
      .map(r => [r.studentId, r.status, r.dorm, r.startDate || getAcademicYearStart_()]);

    const clearRows = Math.max(996, sh.getMaxRows() - 4);
    sh.getRange(5, 1, clearRows, 8).clearContent();

    if (rows.length) {
      sh.getRange(5, 1, rows.length, 4).setValues(rows);
      applyDormRowFormulas_(sh, rows.length);
    }
  });

  applyRegistryStatsFormulas_();
}

/**
 * Writes the Events Attended / Eligible Events / Attendance % / Last Attended
 * formulas for the first numRows data rows (starting at row 5) of a dorm sheet.
 *
 * BUGFIX (v4): Events Attended used to be floored only by the academic year
 * start (Settings!$B$3), while Eligible Events was floored by
 * MAX(start date, academic year start). If a student's own start date was
 * later than the academic year start, an event they attended before their
 * (re-registered) start date still counted as "attended" but no longer
 * counted as "eligible" - producing Attendance % readings over 100%. Both
 * columns now use the same MAX() floor, and Eligible Events also excludes
 * events that have not happened yet (future-dated events no longer drag the
 * percentage down before they occur).
 *
 * v6: Both formulas also now check the event's "Applies To" dorm scope
 * (EventLog!F / AttendanceLog!I) against the row's own dorm ($C). A blank
 * scope means "every dorm"; otherwise the student's dorm must appear in the
 * comma-separated list. The same scope check is applied to BOTH Events
 * Attended and Eligible Events (not just Eligible) so an out-of-scope event
 * never counts toward attendance without also counting toward eligibility -
 * that symmetry is what keeps Attendance % from exceeding 100%, same
 * principle as the v4 fix above.
 *
 * BUGFIX (v6.1): Last Attended did not have the dorm-scope filter applied at
 * all, so it could show the date of an event that wasn't even in scope for
 * that student (e.g. a Dorm D student's Last Attended showing a Dorm-B-only
 * event date) - inconsistent with what Events Attended / Eligible Events
 * were actually counting. It now uses the same scope filter as the other two.
 *
 * BUGFIX (v6.2): Events Attended / Eligible Events used
 * IFERROR(COUNTUNIQUE(FILTER(...)),0) to handle "nothing matches". That's
 * wrong: when FILTER finds zero rows it returns a #N/A error, and Google
 * Sheets' COUNTUNIQUE treats a single error VALUE as one countable "unique
 * value" instead of propagating the error - so COUNTUNIQUE(#N/A) is 1, not an
 * error, and the outer IFERROR never fires. Every student with zero real
 * attendance was showing Events Attended = 1 (and, once any event existed,
 * Eligible Events = 1 too) as soon as they were added to the registry.
 * MAX() does not have this quirk (it correctly propagates the error, which is
 * why Last Attended was already showing blank correctly). Rebuilt using
 * SUMPRODUCT over bounded ranges instead of FILTER/COUNTUNIQUE - it sums a
 * product of 0/1 boolean arrays, so "nothing matches" is a genuine 0, never
 * an error to accidentally swallow.
 */
function applyDormRowFormulas_(sh, numRows) {
  if (!numRows) return;
  const formulas = [];
  for (let i = 0; i < numRows; i++) {
    const row = 5 + i;
    formulas.push([
      '=IF($A' + row + '="","",SUMPRODUCT((AttendanceLog!$E$2:$E$5000=$A' + row + ')*(AttendanceLog!$C$2:$C$5000>=MAX($D' + row + ',Settings!$B$3))*((AttendanceLog!$I$2:$I$5000="")+(IFERROR(SEARCH(","&$C' + row + '&",",","&AttendanceLog!$I$2:$I$5000&","),0)>0))))',
      eligibleEventsFormula_(row, 'D'),
      '=IFERROR(E' + row + '/F' + row + ',0)',
      '=IFERROR(MAX(FILTER(AttendanceLog!$C:$C,AttendanceLog!$E:$E=$A' + row + ',(AttendanceLog!$I:$I="")+(IFERROR(SEARCH(","&$C' + row + '&",",","&AttendanceLog!$I:$I&","),0)>0))),"")'
    ]);
  }
  sh.getRange(5, 5, formulas.length, 4).setFormulas(formulas);
  sh.getRange(5, 4, numRows, 1).setNumberFormat('yyyy-mm-dd');
  sh.getRange(5, 7, numRows, 1).setNumberFormat('0.0%');
  sh.getRange(5, 8, numRows, 1).setNumberFormat('yyyy-mm-dd');
}

/**
 * v7.0: the Eligible Events formula, shared by the dorm sheets (start date in
 * column D) and StudentRegistry (start date in column E). Student ID is
 * always column A and dorm is always column C on both.
 *
 * Two parts added together:
 *  1. Regular events - same as v6.2: every EventLog event on/after the
 *     student's start (and the academic-year start), not in the future, in
 *     their dorm scope - but now EXCLUDING invite-only events (EventLog!G).
 *  2. Invite-only events - one per AttendanceLog row for this student whose
 *     Event Type (AttendanceLog!J) is invite-only, using the same date floor
 *     as Events Attended. AttendanceLog only ever holds ticked invitees, so
 *     an invite-only event becomes "eligible" exactly when it's "attended" -
 *     Attendance % can't exceed 100% and no-shows aren't penalised.
 */
function eligibleEventsFormula_(row, startCol) {
  const floor = 'MAX($' + startCol + row + ',Settings!$B$3)';
  const t = '"' + INVITE_ONLY_TYPE + '"';
  const regular =
    'SUMPRODUCT((EventLog!$C$2:$C$5000>=' + floor + ')' +
    '*(EventLog!$C$2:$C$5000<=TODAY())' +
    '*(EventLog!$G$2:$G$5000<>' + t + ')' +
    '*((EventLog!$F$2:$F$5000="")+(IFERROR(SEARCH(","&$C' + row + '&",",","&EventLog!$F$2:$F$5000&","),0)>0)))';
  const inviteOnly =
    'SUMPRODUCT((AttendanceLog!$E$2:$E$5000=$A' + row + ')' +
    '*(AttendanceLog!$C$2:$C$5000>=' + floor + ')' +
    '*(AttendanceLog!$J$2:$J$5000=' + t + '))';
  return '=IF($A' + row + '="","",' + regular + '+' + inviteOnly + ')';
}

/**
 * Writes the same Events Attended / Eligible Events / Attendance % / Last
 * Attended formulas onto StudentRegistry itself (columns G:J), using each
 * student's own row instead of a per-dorm copy. Unlike the dorm sheets,
 * this covers every student - active or inactive, in any dorm - so it's the
 * one place to see attendance for the whole roster at once. Same dorm-scope
 * filtering and SUMPRODUCT-based counting as applyDormRowFormulas_ (see the
 * comments there, including the v6.2 COUNTUNIQUE/FILTER fix).
 */
function applyRegistryStatsFormulas_() {
  const sh = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(REGISTRY_SHEET);
  if (!sh) return;
  const lastRow = sh.getLastRow();
  if (lastRow < 2) return;
  const numRows = lastRow - 1;
  const formulas = [];
  for (let i = 0; i < numRows; i++) {
    const row = 2 + i;
    formulas.push([
      '=IF($A' + row + '="","",SUMPRODUCT((AttendanceLog!$E$2:$E$5000=$A' + row + ')*(AttendanceLog!$C$2:$C$5000>=MAX($E' + row + ',Settings!$B$3))*((AttendanceLog!$I$2:$I$5000="")+(IFERROR(SEARCH(","&$C' + row + '&",",","&AttendanceLog!$I$2:$I$5000&","),0)>0))))',
      eligibleEventsFormula_(row, 'E'),
      '=IFERROR(G' + row + '/H' + row + ',0)',
      '=IFERROR(MAX(FILTER(AttendanceLog!$C:$C,AttendanceLog!$E:$E=$A' + row + ',(AttendanceLog!$I:$I="")+(IFERROR(SEARCH(","&$C' + row + '&",",","&AttendanceLog!$I:$I&","),0)>0))),"")'
    ]);
  }
  sh.getRange(2, 7, formulas.length, 4).setFormulas(formulas);
  sh.getRange(2, 9, numRows, 1).setNumberFormat('0.0%');
  sh.getRange(2, 10, numRows, 1).setNumberFormat('yyyy-mm-dd');
}

/** Runs a set of consistency checks across the registry, event log, and event sheets and writes results to the DataAudit tab. */
function runDataAudit() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const issues = [];

  // 1. StudentRegistry checks.
  const regSh = ss.getSheetByName(REGISTRY_SHEET);
  const activeDorms = getActiveDormNames_();
  const inactiveDorms = getInactiveDormNames_();
  const allDorms = activeDorms.concat(inactiveDorms);
  const seenIds = {};
  if (regSh && regSh.getLastRow() > 1) {
    const values = regSh.getRange(2, 1, regSh.getLastRow() - 1, 6).getValues();
    values.forEach((r, i) => {
      const row = i + 2;
      const id = normalizeId_(r[0]);
      const dorm = String(r[2]).trim();
      const status = String(r[3]).trim();
      if (!id) { issues.push([REGISTRY_SHEET, row, 'Blank Student ID']); return; }
      if (seenIds[id]) {
        issues.push([REGISTRY_SHEET, row, 'Duplicate Student ID "' + id + '" (also row ' + seenIds[id] + ')']);
      }
      seenIds[id] = row;
      if (!dorm) {
        issues.push([REGISTRY_SHEET, row, 'Missing dorm for ' + id]);
      } else if (allDorms.indexOf(dorm) === -1) {
        issues.push([REGISTRY_SHEET, row, 'Unknown dorm "' + dorm + '" for ' + id]);
      }
      if (['Active', 'Inactive'].indexOf(status) === -1) {
        issues.push([REGISTRY_SHEET, row, 'Invalid status "' + status + '" for ' + id + ' (expected Active or Inactive)']);
      }
      if (!(r[4] instanceof Date)) {
        issues.push([REGISTRY_SHEET, row, 'Missing or invalid start date for ' + id]);
      }
    });
  }

  // 2. EventLog checks.
  const evSh = ss.getSheetByName(EVENT_LOG_SHEET);
  const existingSheetNames = {};
  ss.getSheets().forEach(s => existingSheetNames[s.getName()] = true);
  const seenEventIds = {};
  if (evSh && evSh.getLastRow() > 1) {
    const values = evSh.getRange(2, 1, evSh.getLastRow() - 1, 6).getValues();
    values.forEach((r, i) => {
      const row = i + 2;
      const id = String(r[0]).trim();
      const sheetName = String(r[3]).trim();
      if (!id) { issues.push([EVENT_LOG_SHEET, row, 'Blank Event ID']); return; }
      if (seenEventIds[id]) issues.push([EVENT_LOG_SHEET, row, 'Duplicate Event ID "' + id + '"']);
      seenEventIds[id] = row;
      if (!(r[2] instanceof Date)) issues.push([EVENT_LOG_SHEET, row, 'Missing or invalid event date for ' + id]);
      if (sheetName && !existingSheetNames[sheetName]) {
        issues.push([EVENT_LOG_SHEET, row, 'Sheet "' + sheetName + '" no longer exists for ' + id + ' (run Clean Up Deleted Event References)']);
      }
      const scope = String(r[5] || '').trim();
      if (scope) {
        const unknownDorms = scope.split(',').map(s => s.trim()).filter(d => d && allDorms.indexOf(d) === -1);
        if (unknownDorms.length) {
          issues.push([EVENT_LOG_SHEET, row,
            'Event "' + (r[1] || id) + '" applies to "' + unknownDorms.join(', ') + '", which ' +
            (unknownDorms.length > 1 ? 'aren\u2019t' : 'isn\u2019t') + ' a configured dorm - add ' +
            (unknownDorms.length > 1 ? 'them' : 'it') + ' via Dorms -> Add Dorm, or fix a typo on the event sheet\u2019s "Applies To Dorm(s)" field.']);
        }
      }
    });
  }

  // 3. Event sheet attendee checks: duplicates within a sheet, IDs missing from the registry.
  //    v7.0: on invite-only tabs, also flag a past event with nothing ticked
  //    (almost always means someone forgot to tick the boxes).
  const registry = getRegistryMap_();
  const today = new Date();
  ss.getSheets().forEach(sheet => {
    if (!isEventSheet_(sheet)) return;
    const inviteOnly = isInviteOnlySheet_(sheet);
    const lastRow = Math.max(sheet.getLastRow(), ATTENDEE_START_ROW);
    const numRows = lastRow - ATTENDEE_START_ROW + 1;
    const numCols = inviteOnly ? INVITE_ATTENDED_COL - ATTENDEE_COL + 1 : 1;
    const block = sheet.getRange(ATTENDEE_START_ROW, ATTENDEE_COL, numRows, numCols).getValues();
    const seenInSheet = {};
    let invited = 0;
    let ticked = 0;
    block.forEach((r, i) => {
      const id = normalizeId_(r[0]);
      if (!id) return;
      const row = ATTENDEE_START_ROW + i;
      const label = inviteOnly ? 'invitee' : 'attendee';
      if (seenInSheet[id]) issues.push([sheet.getName(), row, 'Duplicate ' + label + ' entry: ' + id]);
      seenInSheet[id] = true;
      if (!registry[id]) issues.push([sheet.getName(), row, (inviteOnly ? 'Invitee' : 'Attendee') + ' not found in registry: ' + id]);
      if (inviteOnly) {
        invited++;
        if (isTicked_(r[INVITE_ATTENDED_COL - ATTENDEE_COL])) ticked++;
      }
    });
    if (inviteOnly) {
      const eventDate = sheet.getRange('B4').getValue();
      if (invited && !ticked && eventDate instanceof Date && eventDate <= today) {
        issues.push([sheet.getName(), ATTENDEE_START_ROW,
          'Invite-only event with ' + invited + ' invitee(s) but no "Attended?" boxes ticked - nobody is being counted for it yet. Tick column B for everyone who came.']);
      }
    }
  });

  // 4. Attendance recorded before the student's own registered Start Date.
  // This doesn't corrupt anything, but it silently zeroes out Events
  // Attended for that record (the Events Attended / Eligible Events
  // formulas never count anything before a student's Start Date), which
  // looks like "attendance isn't updating" without an obvious cause.
  const attSh = ss.getSheetByName(ATTENDANCE_LOG_SHEET);
  if (attSh && attSh.getLastRow() > 1) {
    const attData = attSh.getRange(2, 1, attSh.getLastRow() - 1, 5).getValues();
    attData.forEach((r, i) => {
      const row = i + 2;
      const eventDate = r[2];
      const studentId = normalizeId_(r[4]);
      const reg = registry[studentId];
      if (!reg || !(eventDate instanceof Date) || !(reg.startDate instanceof Date)) return;
      if (eventDate < reg.startDate) {
        issues.push([ATTENDANCE_LOG_SHEET, row,
          studentId + ' attended on ' + Utilities.formatDate(eventDate, Session.getScriptTimeZone(), 'yyyy-MM-dd') +
          ' but their registered Start Date is ' + Utilities.formatDate(reg.startDate, Session.getScriptTimeZone(), 'yyyy-MM-dd') +
          ' - this attendance won\u2019t count until the Start Date is corrected (Students -> Register Student, same dorm, enter the correct date).']);
      }
    });
  }

  // Write the report.
  let reportSh = ss.getSheetByName(DATA_AUDIT_SHEET);
  if (!reportSh) reportSh = ss.insertSheet(DATA_AUDIT_SHEET);
  reportSh.showSheet();
  reportSh.clear();
  reportSh.getRange('A1:B1').setValues([['Audit Run At', new Date()]]);
  reportSh.getRange('B1').setNumberFormat('yyyy-mm-dd hh:mm');
  reportSh.getRange('A2:B2').setValues([['Total Issues', issues.length]]);
  reportSh.getRange('A4:C4').setValues([['Sheet', 'Row', 'Issue']]);
  styleHeader_(reportSh.getRange('A4:C4'));
  if (issues.length) reportSh.getRange(5, 1, issues.length, 3).setValues(issues);
  reportSh.setFrozenRows(4);
  reportSh.autoResizeColumns(1, 3);
  ss.setActiveSheet(reportSh);

  SpreadsheetApp.getUi().alert('Data audit complete. Found ' + issues.length + ' issue(s). See the "' + DATA_AUDIT_SHEET + '" tab for details.');
}

/**
 * Moves EventLog/AttendanceLog rows dated before the current academic year
 * into archive sheets and hides the underlying event tabs (nothing is
 * deleted). This also frees up the 2-digit Event ID range for the new year,
 * since makeEventId_() only looks at what's still in the live EventLog sheet.
 */
function archivePreviousYearEvents() {
  const ui = SpreadsheetApp.getUi();
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const cutoff = getAcademicYearStart_();

  const confirm = ui.alert(
    'Archive Previous Years\u2019 Events',
    'This moves EventLog / AttendanceLog rows dated before ' +
    Utilities.formatDate(cutoff, Session.getScriptTimeZone(), 'yyyy-MM-dd') +
    ' into archive sheets and hides their event tabs. Nothing is deleted. Continue?',
    ui.ButtonSet.YES_NO
  );
  if (confirm !== ui.Button.YES) return;

  const evSh = ss.getSheetByName(EVENT_LOG_SHEET);
  const attSh = ss.getSheetByName(ATTENDANCE_LOG_SHEET);
  const evArchive = ensureSheet_(EVENT_LOG_ARCHIVE_SHEET);
  const attArchive = ensureSheet_(ATTENDANCE_LOG_ARCHIVE_SHEET);

  ensureEventArchiveHeader_(evArchive);
  ensureAttendanceArchiveHeader_(attArchive);

  let archivedEvents = 0;
  let archivedAttendance = 0;
  const archivedEventIds = {};
  const archivedSheetNames = {};

  if (evSh && evSh.getLastRow() > 1) {
    const data = evSh.getRange(2, 1, evSh.getLastRow() - 1, EVENT_LOG_COLS).getValues();
    const keep = [];
    const toArchive = [];
    data.forEach(r => {
      if (r[2] instanceof Date && r[2] < cutoff) {
        toArchive.push(r);
        archivedEventIds[String(r[0]).trim()] = true;
        if (r[3]) archivedSheetNames[String(r[3]).trim()] = true;
      } else {
        keep.push(r);
      }
    });
    if (toArchive.length) {
      evArchive.getRange(evArchive.getLastRow() + 1, 1, toArchive.length, EVENT_LOG_COLS).setValues(toArchive);
      archivedEvents = toArchive.length;
      evSh.getRange(2, 1, evSh.getMaxRows() - 1, EVENT_LOG_COLS).clearContent();
      if (keep.length) evSh.getRange(2, 1, keep.length, EVENT_LOG_COLS).setValues(keep);
    }
  }

  if (attSh && attSh.getLastRow() > 1 && Object.keys(archivedEventIds).length) {
    const data = attSh.getRange(2, 1, attSh.getLastRow() - 1, ATTENDANCE_LOG_COLS).getValues();
    const keep = [];
    const toArchive = [];
    data.forEach(r => {
      if (archivedEventIds[String(r[1]).trim()]) toArchive.push(r);
      else keep.push(r);
    });
    if (toArchive.length) {
      attArchive.getRange(attArchive.getLastRow() + 1, 1, toArchive.length, ATTENDANCE_LOG_COLS).setValues(toArchive);
      archivedAttendance = toArchive.length;
      attSh.getRange(2, 1, attSh.getMaxRows() - 1, ATTENDANCE_LOG_COLS).clearContent();
      if (keep.length) attSh.getRange(2, 1, keep.length, ATTENDANCE_LOG_COLS).setValues(keep);
    }
  }

  // Hide (never delete) the underlying event sheets that were archived.
  Object.keys(archivedSheetNames).forEach(name => {
    const sh = ss.getSheetByName(name);
    if (sh && !sh.isSheetHidden()) sh.hideSheet();
  });

  [EVENT_LOG_ARCHIVE_SHEET, ATTENDANCE_LOG_ARCHIVE_SHEET].forEach(name => {
    const sh = ss.getSheetByName(name);
    if (sh && !sh.isSheetHidden()) sh.hideSheet();
  });

  rebuildDormSummarySheets();
  ui.alert(
    'Archived ' + archivedEvents + ' event(s) and ' + archivedAttendance + ' attendance record(s).\n\n' +
    'Their event tabs were hidden, not deleted - unhide them from the sheet tab bar (right-click a visible tab -> "Show hidden sheets") if you need to look something up.'
  );
}

// v6.9: handleSpreadsheetChange() and installTriggers_() were removed. They
// existed only to set up installable triggers, which is what required the
// "Allow this application to run when you are not present" permission.
// Deleted-tab cleanup is now handled by onOpen(), by every attendance
// refresh (which rebuilds the logs from existing tabs), and by the manual
// Events -> Clean Up Deleted Event References menu item below.

function cleanUpDeletedEventReferences() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const before = ss.getSheetByName(EVENT_LOG_SHEET);
  const beforeCount = before ? Math.max(before.getLastRow() - 1, 0) : 0;
  pruneOrphanedEventLogEntries_();
  const after = ss.getSheetByName(EVENT_LOG_SHEET);
  const afterCount = after ? Math.max(after.getLastRow() - 1, 0) : 0;
  SpreadsheetApp.getUi().alert('Removed ' + (beforeCount - afterCount) + ' orphaned EventLog entrie(s) whose sheet no longer exists.');
}

/** Removes EventLog/AttendanceLog rows whose "Sheet Name" no longer matches an existing tab. Safe to run any time. */
function pruneOrphanedEventLogEntries_() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const evSh = ss.getSheetByName(EVENT_LOG_SHEET);
  if (!evSh || evSh.getLastRow() < 2) return;

  const existingSheetNames = {};
  ss.getSheets().forEach(s => existingSheetNames[s.getName()] = true);

  const data = evSh.getRange(2, 1, evSh.getLastRow() - 1, EVENT_LOG_COLS).getValues();
  const keptRows = [];
  const removedEventIds = {};
  data.forEach(r => {
    const sheetName = String(r[3]).trim();
    if (sheetName && !existingSheetNames[sheetName]) {
      removedEventIds[String(r[0]).trim()] = true;
    } else {
      keptRows.push(r);
    }
  });

  if (!Object.keys(removedEventIds).length) return;

  evSh.getRange(2, 1, evSh.getMaxRows() - 1, EVENT_LOG_COLS).clearContent();
  if (keptRows.length) evSh.getRange(2, 1, keptRows.length, EVENT_LOG_COLS).setValues(keptRows);

  const attSh = ss.getSheetByName(ATTENDANCE_LOG_SHEET);
  if (attSh && attSh.getLastRow() > 1) {
    const attData = attSh.getRange(2, 1, attSh.getLastRow() - 1, ATTENDANCE_LOG_COLS).getValues();
    const keptAtt = attData.filter(r => !removedEventIds[String(r[1]).trim()]);
    attSh.getRange(2, 1, attSh.getMaxRows() - 1, ATTENDANCE_LOG_COLS).clearContent();
    if (keptAtt.length) attSh.getRange(2, 1, keptAtt.length, ATTENDANCE_LOG_COLS).setValues(keptAtt);
  }

  rebuildDormSummarySheets();
}

/** Locks backend sheets to menu-only edits and warns before formula columns get overwritten by hand. Safe to re-run. */
function protectSystemSheets() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();

  [REGISTRY_SHEET, EVENT_LOG_SHEET, ATTENDANCE_LOG_SHEET, SETTINGS_SHEET, TEMPLATE_SHEET, TRANSFER_LOG_SHEET, DELETED_STUDENTS_SHEET].forEach(name => {
    const sh = ss.getSheetByName(name);
    if (!sh) return;
    sh.getProtections(SpreadsheetApp.ProtectionType.SHEET).forEach(p => p.remove());
    const protection = sh.protect().setDescription('System sheet - edit only via the Dorm Attendance menu');
    try {
      if (protection.canDomainEdit()) protection.setDomainEdit(false);
    } catch (err) {
      // Personal Google accounts don't have a domain concept; ignore.
    }
  });

  ss.getSheets().forEach(sheet => {
    if (!isEventSheet_(sheet)) return;
    protectFormulaColumns_(sheet, sheet.getRange(ATTENDEE_START_ROW, 3, 993, 2)); // C:D lookup formulas
  });

  getActiveDormNames_().forEach(dorm => {
    const sh = ss.getSheetByName(dorm);
    if (!sh) return;
    protectFormulaColumns_(sh, sh.getRange(5, 5, Math.max(sh.getMaxRows() - 4, 1), 4)); // E:H computed columns
  });

  SpreadsheetApp.getUi().alert(
    'Protection applied.\n\n' +
    '\u2022 StudentRegistry, EventLog, AttendanceLog, Settings, EventTemplate, TransferLog, and DeletedStudents are now locked to menu-only edits.\n' +
    '\u2022 Formula columns on event sheets and dorm sheets now warn before anyone overwrites them by hand.\n\n' +
    'To let other staff edit the locked sheets directly, go to Data -> Protected sheets and ranges, open each entry, and add their email under "Who can edit".'
  );
}

function protectFormulaColumns_(sheet, range) {
  const a1 = range.getA1Notation();
  sheet.getProtections(SpreadsheetApp.ProtectionType.RANGE)
    .filter(p => p.getRange().getA1Notation() === a1)
    .forEach(p => p.remove());
  const protection = range.protect().setDescription('Formula column - do not overwrite');
  protection.setWarningOnly(true);
}

/** Moves an existing, active student to a different active dorm while keeping their Start Date (and therefore their attendance history) intact. */
function transferStudentDorm() {
  const ui = SpreadsheetApp.getUi();

  const idResponse = ui.prompt('Transfer Student', 'Enter the student ID number:', ui.ButtonSet.OK_CANCEL);
  if (idResponse.getSelectedButton() !== ui.Button.OK) return;
  const studentId = normalizeId_(idResponse.getResponseText());
  if (!studentId) {
    ui.alert('Student ID is required.');
    return;
  }

  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sh = ss.getSheetByName(REGISTRY_SHEET);
  const row = findRegistryRow_(studentId);
  if (!row) {
    ui.alert('Student ID was not found in the registry. Use Students -> Register Student for a brand-new student.');
    return;
  }

  const currentDorm = String(sh.getRange(row, 3).getValue()).trim();
  const activeDorms = getActiveDormNames_();

  const dormResponse = ui.prompt(
    'Transfer Student',
    'Currently in: ' + currentDorm + '.\nEnter the NEW dorm exactly as one of:\n' + activeDorms.join(', '),
    ui.ButtonSet.OK_CANCEL
  );
  if (dormResponse.getSelectedButton() !== ui.Button.OK) return;
  const newDorm = dormResponse.getResponseText().trim();
  if (activeDorms.indexOf(newDorm) === -1) {
    ui.alert('Invalid dorm. Please use one of the active dorms: ' + activeDorms.join(', '));
    return;
  }
  if (newDorm === currentDorm) {
    ui.alert('Student is already in that dorm. No changes made.');
    return;
  }

  const confirm = ui.alert(
    'Transfer Student',
    'Move ' + studentId + ' from ' + currentDorm + ' to ' + newDorm + '?\n\n' +
    'Their attendance history and original start date carry over - they are NOT reset.',
    ui.ButtonSet.YES_NO
  );
  if (confirm !== ui.Button.YES) return;

  const now = new Date();
  sh.getRange(row, 3).setValue(newDorm);   // Dorm - Start Date (column E) is intentionally left untouched.
  sh.getRange(row, 4).setValue('Active');  // A transfer implies the student is active in the new dorm.
  sh.getRange(row, 6).setValue(now);       // Last Updated

  logDormTransfer_(studentId, currentDorm, newDorm, now);
  rebuildDormSummarySheets();
  ui.alert('Transferred ' + studentId + ' from ' + currentDorm + ' to ' + newDorm + '. Attendance history carried over.');
}

function logDormTransfer_(studentId, fromDorm, toDorm, when) {
  const sh = ensureSheet_(TRANSFER_LOG_SHEET);
  setupTransferLog_();
  sh.appendRow([studentId, fromDorm, toDorm, when]);
}

function openBulkRegisterDialog() {
  const html = HtmlService.createHtmlOutputFromFile('BulkRegister')
    .setWidth(600)
    .setHeight(500);
  SpreadsheetApp.getUi().showModalDialog(html, 'Bulk Register Students');
}

function processBulkRegistration(tsvData) {
  if (!tsvData || !tsvData.trim()) return 'No data provided.';

  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sh = ss.getSheetByName(REGISTRY_SHEET);
  const activeDorms = getActiveDormNames_();
  const now = new Date();

  // Excel copy-paste splits rows by newlines (\n) and columns by tabs (\t).
  const rows = tsvData.trim().split('\n');
  let added = 0;
  let updated = 0;
  const errors = [];
  const seenInPaste = {};

  rows.forEach((row, index) => {
    if (!row.trim()) return;

    const cols = row.split('\t');
    const studentId = normalizeId_(cols[0]);
    if (!studentId) return; // Skip completely empty rows

    if (seenInPaste[studentId]) {
      errors.push(`Row ${index + 1} (ID: ${studentId}): duplicate entry in this paste - later row wins`);
    }
    seenInPaste[studentId] = true;

    const dorm = String(cols[1] || '').trim();
    if (activeDorms.indexOf(dorm) === -1) {
      errors.push(`Row ${index + 1} (ID: ${studentId}): Invalid dorm '${dorm}'`);
      return; // Skip this student, move to the next
    }

    const startText = String(cols[2] || '').trim();
    const notes = String(cols[3] || '').trim();
    const existingRow = findRegistryRow_(studentId);

    // BUGFIX (v4): re-pasting an existing student used to always reset their
    // Start Date to "now" when the Start Date column was left blank, which
    // silently wiped their attendance history eligibility. Now a blank Start
    // Date keeps whatever is already on file for an existing student.
    let startDate;
    if (startText) {
      startDate = parseDate_(startText);
      if (!startDate) {
        errors.push(`Row ${index + 1} (ID: ${studentId}): Invalid start date '${startText}'`);
        return;
      }
    } else if (existingRow) {
      const existingStart = sh.getRange(existingRow, 5).getValue();
      startDate = existingStart instanceof Date ? existingStart : now;
    } else {
      startDate = now;
    }

    if (existingRow) {
      sh.getRange(existingRow, 2, 1, 5).setValues([[notes, dorm, 'Active', startDate, now]]);
      updated++;
    } else {
      sh.appendRow([studentId, notes, dorm, 'Active', startDate, now]);
      added++;
    }
  });

  if (added > 0 || updated > 0) {
    rebuildDormSummarySheets();
  }

  let resultMsg = `<span style="color:green">Success! Added: ${added}, Updated: ${updated}</span>`;
  if (errors.length > 0) {
    resultMsg += `<br><br><div class="error"><b>Notes:</b><br>${errors.join('<br>')}</div>`;
  }
  return resultMsg;
}

function openBulkDeactivateDialog() {
  const html = HtmlService.createHtmlOutputFromFile('BulkDeactivate')
    .setWidth(600)
    .setHeight(500);
  SpreadsheetApp.getUi().showModalDialog(html, 'Bulk Deactivate Students');
}

function processBulkDeactivation(idsText) {
  if (!idsText || !idsText.trim()) return 'No data provided.';

  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sh = ss.getSheetByName(REGISTRY_SHEET);
  const now = new Date();
  const lines = idsText.trim().split('\n');
  let deactivated = 0;
  let alreadyInactive = 0;
  const notFound = [];

  lines.forEach(line => {
    if (!line.trim()) return;
    const cols = line.split('\t');
    const studentId = normalizeId_(cols[0]);
    if (!studentId) return;

    const row = findRegistryRow_(studentId);
    if (!row) {
      notFound.push(studentId);
      return;
    }

    const currentStatus = String(sh.getRange(row, 4).getValue()).trim();
    if (currentStatus === 'Inactive') {
      alreadyInactive++;
      return;
    }

    // Status only - the student's row (and history) stays in the sheet.
    sh.getRange(row, 4).setValue('Inactive');
    sh.getRange(row, 6).setValue(now);
    deactivated++;
  });

  if (deactivated > 0) rebuildDormSummarySheets();

  let resultMsg = `<span style="color:green">Deactivated: ${deactivated}. Already inactive: ${alreadyInactive}.</span>`;
  if (notFound.length) {
    resultMsg += `<br><br><div class="error"><b>Not found (skipped):</b><br>${notFound.join('<br>')}</div>`;
  }
  return resultMsg;
}

/**
 * Entry point called from ImportEvent.html.
 *
 * v4.1: the dialog now expects a structured table copied straight out of
 * Excel/Sheets - a header/data row with Event Name, Event Date, Venue,
 * Attendee Count, Expected Attendee Count, followed by an "Attendee ID" /
 * "Attendee Dorm" table. That structured format is tried first; if no
 * "Attendee ID" row is found at all, this falls back to the original
 * freeform extraction (regex-scan for a DD/MM/YYYY date and ID-shaped
 * tokens) for genuinely messy, unstructured pastes.
 */
function processEventImport(eventNameOverride, rawText) {
  if (!rawText || !rawText.trim()) {
    throw new Error('Please paste your event data first.');
  }

  const structured = parseStructuredEventImport_(rawText);
  if (structured) {
    if (eventNameOverride && eventNameOverride.trim()) {
      structured.eventName = eventNameOverride.trim(); // Explicit override wins.
    }
    if (!structured.eventName) {
      throw new Error('Could not find an event name. Add an "Event Name" column to your paste, or fill in the Event Name field above.');
    }
    if (!structured.eventDate) {
      throw new Error('Could not read the Event Date. Use a format like 2026/6/26 or 2026-06-26.');
    }
    return createImportedEventSheet_(structured);
  }

  // No "Attendee ID" table was found - treat this as an old messy paste.
  if (!eventNameOverride || !eventNameOverride.trim()) {
    throw new Error('No "Attendee ID" table was detected in your paste, so an Event Name is required. Fill in the Event Name field above, or paste a structured table (see the format shown above the box).');
  }
  return processMessyEventImport_(eventNameOverride.trim(), rawText);
}

/**
 * Parses a structured, tab-separated paste like:
 *   Event Name  Event Date  Venue      Attendee Count  Expected Attendee Count
 *   Dorm B CCM  2026/6/26   Dorm B 1F  5               10
 *
 *   Attendee ID  Attendee Dorm
 *   2025B638     Dorm B
 * Returns null (so the caller can fall back to messy parsing) if no row
 * starting with "Attendee ID" or "Student ID" is found anywhere.
 */
function parseStructuredEventImport_(rawText) {
  const rows = rawText.replace(/\r/g, '').split('\n').map(line => line.split('\t').map(c => c.trim()));

  const headerIdx = rows.findIndex(cols => (cols[0] || '').toLowerCase() === 'event name');
  let eventName = '';
  let eventDate = null;
  let venue = '';
  let expectedCount = '';
  let reportedCount = '';
  let searchFrom = 0;

  if (headerIdx >= 0) {
    let dataIdx = headerIdx + 1;
    while (dataIdx < rows.length && rows[dataIdx].every(c => !c)) dataIdx++;
    if (dataIdx < rows.length) {
      const dr = rows[dataIdx];
      eventName = dr[0] || '';
      eventDate = parseDate_(dr[1] || '');
      venue = dr[2] || '';
      reportedCount = dr[3] || '';
      expectedCount = dr[4] || '';
      searchFrom = dataIdx + 1;
    }
  }

  let attendeeHeaderIdx = -1;
  for (let i = searchFrom; i < rows.length; i++) {
    const first = (rows[i][0] || '').toLowerCase();
    if (first === 'attendee id' || first === 'student id') {
      attendeeHeaderIdx = i;
      break;
    }
  }
  if (attendeeHeaderIdx === -1) return null;

  const seen = {};
  const attendees = [];
  let duplicates = 0;
  for (let i = attendeeHeaderIdx + 1; i < rows.length; i++) {
    const id = normalizeId_(rows[i][0]);
    if (!id) continue;
    const dorm = rows[i][1] || '';
    if (seen[id]) { duplicates++; continue; }
    seen[id] = true;
    attendees.push({ id: id, dorm: dorm });
  }

  return {
    eventName: eventName,
    eventDate: eventDate,
    venue: venue,
    expectedCount: expectedCount,
    reportedCount: reportedCount,
    attendees: attendees,
    duplicates: duplicates
  };
}

/** Original freeform extraction for messy, unstructured old paper records - no "Attendee ID" table required. */
function processMessyEventImport_(eventName, rawText) {
  // 1. Extract the Date (looks for DD/MM/YYYY)
  let eventDate = new Date();
  const dateMatch = rawText.match(/(\d{1,2})\/(\d{1,2})\/(\d{4})/);
  if (dateMatch) {
    // JavaScript dates format as (Year, MonthIndex, Day). Month is 0-indexed.
    eventDate = new Date(dateMatch[3], dateMatch[2] - 1, dateMatch[1]);
  }

  // 2. Extract the Student IDs
  const idRegex = /[a-zA-Z0-9]*\d[a-zA-Z0-9]*/gi;
  const rawMatches = rawText.match(idRegex) || [];

  // Filter out random short numbers (like the attendee count '14' or the year '2026')
  // We assume a real Student ID is at least 5 characters long
  const validMatches = rawMatches.filter(id => id.length >= 5);

  // Remove duplicates and capitalize all letters
  const uniqueIds = [...new Set(validMatches)].map(id => id.toUpperCase());

  if (uniqueIds.length === 0) {
    throw new Error('Could not find any valid Student IDs in the pasted text. Check the formatting, or use the structured table format shown above the box.');
  }

  return createImportedEventSheet_({
    eventName: eventName,
    eventDate: eventDate,
    venue: '',
    expectedCount: '',
    reportedCount: '',
    attendees: uniqueIds.map(id => ({ id: id, dorm: '' })),
    duplicates: 0
  });
}

/** Shared sheet-creation step for both the structured and messy import paths. */
function createImportedEventSheet_(data) {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const eventId = makeEventId_(data.eventDate);
  const sheetNameBase = EVENT_PREFIX + Utilities.formatDate(data.eventDate, Session.getScriptTimeZone(), 'yyyyMMdd') + '_' + sanitizeSheetName_(data.eventName);
  const sheetName = uniqueSheetName_(sheetNameBase.substring(0, 90));
  const dormScope = detectDormScope_(data.eventName);

  let template = ss.getSheetByName(TEMPLATE_SHEET);
  if (!template) {
    setupEventTemplate_();
    template = ss.getSheetByName(TEMPLATE_SHEET);
  }

  const newSheet = template.copyTo(ss).setName(sheetName);
  ss.setActiveSheet(newSheet);
  ss.moveActiveSheet(ss.getNumSheets());

  newSheet.getRange('B3').setValue(eventId);
  newSheet.getRange('B4').setValue(data.eventDate).setNumberFormat('yyyy-mm-dd');
  newSheet.getRange('B5').setValue(data.eventName);
  if (data.venue) newSheet.getRange('E3').setValue(data.venue);
  if (data.expectedCount !== '') newSheet.getRange('E4').setValue(isNaN(Number(data.expectedCount)) ? data.expectedCount : Number(data.expectedCount));
  if (data.reportedCount !== '') newSheet.getRange('E5').setValue(isNaN(Number(data.reportedCount)) ? data.reportedCount : Number(data.reportedCount));
  newSheet.getRange('E6').setValue(dormScope);
  applyEventSheetFormatting_(newSheet);

  // Paste attendee IDs starting at A8, with any pasted dorm noted in column B for reference.
  if (data.attendees.length) {
    const idCol = data.attendees.map(a => [a.id]);
    const notesCol = data.attendees.map(a => [a.dorm ? 'Imported dorm: ' + a.dorm : '']);
    newSheet.getRange(ATTENDEE_START_ROW, ATTENDEE_COL, idCol.length, 1).setValues(idCol);
    newSheet.getRange(ATTENDEE_START_ROW, ATTENDEE_COL + 1, notesCol.length, 1).setValues(notesCol);
  }

  upsertEventLog_(eventId, data.eventName, data.eventDate, sheetName, dormScope);

  // Sync immediately - the onEdit auto-refresh never fires for edits made by
  // the script itself, so without this a freshly imported event's attendance
  // wouldn't show up on the dorm sheet / StudentRegistry until someone
  // remembered Refresh Attendance Log.
  const refreshResult = refreshAttendanceLogSilently_();
  const sheetPrefix = sheetName + ': unregistered ID skipped: ';
  const skippedIds = refreshResult.warnings
    .filter(w => w.indexOf(sheetPrefix) === 0)
    .map(w => w.slice(sheetPrefix.length));

  let msg = `Success! Created sheet '${sheetName}' (Event ID ${eventId}) and logged ${data.attendees.length} attendee(s).`;
  msg += ` Applies to: ${dormScope ? dormScope.split(',').join(', ') : 'All dorms'} (edit cell E6 on the sheet to change this).`;
  if (data.duplicates) {
    msg += ` (${data.duplicates} duplicate attendee row(s) were skipped.)`;
  }
  if (data.reportedCount !== '' && !isNaN(Number(data.reportedCount)) && Number(data.reportedCount) !== data.attendees.length) {
    msg += `<br><br><span style="color:#cc7a00">Note: the pasted "Attendee Count" said ${data.reportedCount}, but ${data.attendees.length} attendee row(s) were found below the Attendee ID header. Double-check the paste if that's unexpected.</span>`;
  }
  if (skippedIds.length) {
    msg += `<br><br><span style="color:#cc7a00">Not yet counted toward attendance (register them first, then run Refresh Attendance Log): ${skippedIds.join(', ')}</span>`;
  }
  return msg;
}

function openImportEventDialog() {
  const html = HtmlService.createHtmlOutputFromFile('ImportEvent')
    .setWidth(600)
    .setHeight(500);
  SpreadsheetApp.getUi().showModalDialog(html, 'Import Old Event Records');
}

// ---------------------------------------------------------------------------
// v7.0: Invite-only events
// ---------------------------------------------------------------------------

function openImportInviteOnlyDialog() {
  const html = HtmlService.createHtmlOutputFromFile('ImportInviteOnlyEvent')
    .setWidth(640)
    .setHeight(600);
  SpreadsheetApp.getUi().showModalDialog(html, 'Import Invite-Only Event');
}

/**
 * Entry point called from ImportInviteOnlyEvent.html.
 *
 * Expected paste (copied from BulkActionTemplates section 4, headers included):
 *   Event Name   Event Date   Venue
 *   Sushi Night  2026-07-11   Dorm A Lounge
 *
 *   Invited ID   Dorm         Attended?
 *   2025B638     Dorm B       TRUE
 *   2025B639     Dorm B       FALSE
 *
 * The event-details block is optional if Event Name / Event Date are typed
 * into the dialog instead (those boxes win when filled). The Dorm and
 * Attended? columns are optional too - a bare list of IDs works, and
 * everyone starts unticked so staff can tick on the event tab afterwards.
 */
function processInviteOnlyEventImport(eventNameOverride, eventDateOverride, rawText) {
  if (!rawText || !rawText.trim()) {
    throw new Error('Please paste the invite list first.');
  }

  const data = parseInviteOnlyImport_(rawText);

  if (eventNameOverride && String(eventNameOverride).trim()) {
    data.eventName = String(eventNameOverride).trim();
  }
  if (eventDateOverride && String(eventDateOverride).trim()) {
    const d = parseDate_(String(eventDateOverride).trim());
    if (!d) throw new Error('Could not read the Event Date you typed. Use yyyy-mm-dd, e.g. 2026-07-11.');
    data.eventDate = d;
  }

  if (!data.eventName) {
    throw new Error('Could not find an event name. Include the "Event Name" header row in your paste, or type it into the Event Name box.');
  }
  if (!data.eventDate) {
    throw new Error('Could not read the Event Date. Include the "Event Date" column in your paste (e.g. 2026-07-11), or pick it in the Event Date box.');
  }
  if (!data.invitees.length) {
    throw new Error('No invited student IDs were found. Make sure the paste includes the "Invited ID" header row and the IDs below it.');
  }

  return createInviteOnlyEventSheet_(data);
}

/**
 * Parses an invite-only paste. Returns
 *   { eventName, eventDate, venue, invitees: [{id, dorm, attended}], duplicates, hasAttendedColumn }
 * Never returns null - with no recognisable "Invited ID" header it falls back
 * to treating every ID-looking first cell as an invitee.
 */
function parseInviteOnlyImport_(rawText) {
  const rows = rawText.replace(/\r/g, '').split('\n').map(line => line.split('\t').map(c => c.trim()));

  // 1. Optional event-details block ("Event Name | Event Date | Venue" + one data row).
  let eventName = '';
  let eventDate = null;
  let venue = '';
  let searchFrom = 0;
  const headerIdx = rows.findIndex(cols => (cols[0] || '').toLowerCase() === 'event name');
  if (headerIdx >= 0) {
    const headers = rows[headerIdx].map(h => h.toLowerCase());
    const dateCol = Math.max(headers.indexOf('event date'), 1);
    const venueCol = headers.indexOf('venue') >= 0 ? headers.indexOf('venue') : 2;
    let dataIdx = headerIdx + 1;
    while (dataIdx < rows.length && rows[dataIdx].every(c => !c)) dataIdx++;
    if (dataIdx < rows.length) {
      const dr = rows[dataIdx];
      eventName = dr[0] || '';
      eventDate = parseDate_(dr[dateCol] || '');
      venue = dr[venueCol] || '';
      searchFrom = dataIdx + 1;
    }
  }

  // 2. Invite list. Find its header row: first cell like "Invited ID",
  //    "Invitee ID", "Attendee ID", "Student ID" or "Resident ID".
  const idHeaderRe = /^(invited|invitee|attendee|student|resident)\s*id$/i;
  let listHeaderIdx = -1;
  for (let i = searchFrom; i < rows.length; i++) {
    if (idHeaderRe.test(rows[i][0] || '')) { listHeaderIdx = i; break; }
  }

  const byId = {};
  const order = [];
  let duplicates = 0;
  let hasAttendedColumn = false;
  const addInvitee = (id, dorm, attended) => {
    if (byId[id]) {
      duplicates++;
      byId[id].attended = byId[id].attended || attended; // ticked anywhere = ticked
      if (!byId[id].dorm && dorm) byId[id].dorm = dorm;
      return;
    }
    byId[id] = { id: id, dorm: dorm, attended: attended };
    order.push(id);
  };

  if (listHeaderIdx >= 0) {
    const headers = rows[listHeaderIdx];
    let dormCol = -1;
    let attCol = -1;
    headers.forEach((h, c) => {
      if (c === 0 || !h) return;
      if (dormCol === -1 && /dorm/i.test(h)) { dormCol = c; return; }
      if (attCol === -1 && /attend|came|present|show|tick|check|\u2713|\u2714/i.test(h)) attCol = c;
    });
    hasAttendedColumn = attCol >= 0;

    for (let i = listHeaderIdx + 1; i < rows.length; i++) {
      const id = normalizeId_(rows[i][0]);
      if (!id || idHeaderRe.test(id)) continue;
      const dorm = dormCol >= 0 ? (rows[i][dormCol] || '') : '';
      const attended = attCol >= 0 ? isTicked_(rows[i][attCol]) : false;
      addInvitee(id, dorm, attended);
    }
  } else {
    // Fallback: no header row - any line whose first cell looks like a
    // student ID (contains a digit, 5+ characters, no spaces) is an invitee.
    // Any other cell that reads like a tick (TRUE / Yes / x / ✓) marks them
    // attended; any cell matching a dorm name is kept as their pasted dorm.
    const dormNames = getActiveDormNames_().concat(getInactiveDormNames_()).map(d => d.toLowerCase());
    for (let i = searchFrom; i < rows.length; i++) {
      const cells = rows[i];
      const first = cells[0] || '';
      if (!/\d/.test(first) || first.length < 5 || /\s/.test(first)) continue;
      if (/^\d{1,4}[-\/.]\d{1,2}[-\/.]\d{1,4}$/.test(first)) continue; // a date, not an ID
      const rest = cells.slice(1);
      const dorm = rest.find(c => dormNames.indexOf(c.toLowerCase()) >= 0) || '';
      const attended = rest.some(c => isTicked_(c));
      if (attended) hasAttendedColumn = true;
      addInvitee(normalizeId_(first), dorm, attended);
    }
  }

  return {
    eventName: eventName,
    eventDate: eventDate,
    venue: venue,
    invitees: order.map(id => byId[id]),
    duplicates: duplicates,
    hasAttendedColumn: hasAttendedColumn
  };
}

/** Builds the invite-only event tab from parsed data, logs it, refreshes, and returns an HTML status message for the dialog. */
function createInviteOnlyEventSheet_(data) {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const eventId = makeEventId_(data.eventDate);
  const sheetNameBase = EVENT_PREFIX + Utilities.formatDate(data.eventDate, Session.getScriptTimeZone(), 'yyyyMMdd') + '_' + sanitizeSheetName_(data.eventName);
  const sheetName = uniqueSheetName_(sheetNameBase.substring(0, 90));

  let template = ss.getSheetByName(TEMPLATE_SHEET);
  if (!template) {
    setupEventTemplate_();
    template = ss.getSheetByName(TEMPLATE_SHEET);
  }

  const newSheet = template.copyTo(ss).setName(sheetName);
  ss.setActiveSheet(newSheet);
  ss.moveActiveSheet(ss.getNumSheets());

  newSheet.getRange('A1').setValue('Event Attendance Sheet (Invite-Only)');
  newSheet.getRange('B3').setValue(eventId);
  newSheet.getRange('B4').setValue(data.eventDate).setNumberFormat('yyyy-mm-dd');
  newSheet.getRange('B5').setValue(data.eventName);
  if (data.venue) newSheet.getRange('E3').setValue(data.venue);
  applyEventSheetFormatting_(newSheet);
  formatInviteOnlySheet_(newSheet, data.invitees.length);

  if (data.invitees.length) {
    const n = data.invitees.length;
    newSheet.getRange(ATTENDEE_START_ROW, ATTENDEE_COL, n, 1).setValues(data.invitees.map(p => [p.id]));
    newSheet.getRange(ATTENDEE_START_ROW, INVITE_ATTENDED_COL, n, 1).setValues(data.invitees.map(p => [p.attended === true]));
    newSheet.getRange(ATTENDEE_START_ROW, 5, n, 1).setValues(data.invitees.map(p => [p.dorm ? 'Imported dorm: ' + p.dorm : '']));
  }

  upsertEventLog_(eventId, data.eventName, data.eventDate, sheetName, '', INVITE_ONLY_TYPE);

  // Sync immediately - onEdit never fires for the script's own edits.
  refreshAttendanceLogSilently_();

  const registry = getRegistryMap_();
  const ticked = data.invitees.filter(p => p.attended);
  const unregisteredTicked = ticked.filter(p => !registry[p.id]).map(p => p.id);
  const unregisteredUnticked = data.invitees.filter(p => !p.attended && !registry[p.id]).map(p => p.id);
  const countedNow = ticked.length - unregisteredTicked.length;

  let msg = `Success! Created invite-only sheet '${sheetName}' (Event ID ${eventId}).<br>` +
    `Invited: ${data.invitees.length} &nbsp;|&nbsp; Ticked as attended: ${ticked.length} &nbsp;|&nbsp; Counted on the dashboard now: ${countedNow}`;
  if (data.duplicates) {
    msg += `<br>(${data.duplicates} duplicate invitee row(s) were merged.)`;
  }
  msg += '<br><br><span style="color:#333;font-weight:normal">Only ticked residents get +1 Events Attended and +1 Eligible Events. Invited no-shows and everyone who wasn\u2019t invited are unaffected.</span>';
  if (!ticked.length) {
    msg += '<br><br><span style="color:#cc7a00">Nobody is ticked yet. Open the new tab and tick the "Attended?" box (column B) for each resident who came - the dashboard updates as you tick.</span>';
  }
  if (unregisteredTicked.length) {
    msg += `<br><br><span style="color:#cc7a00">Ticked but not in the registry, so not counted yet (register them, then run Refresh Attendance Log): ${unregisteredTicked.join(', ')}</span>`;
  }
  if (unregisteredUnticked.length) {
    msg += `<br><br><span style="color:#cc7a00">Invited but not in the registry (will need registering before a tick counts): ${unregisteredUnticked.join(', ')}</span>`;
  }
  return msg;
}

/** Turns a freshly copied event tab into the invite-only layout (labels, Event Type marker, checkbox column, live counts). */
function formatInviteOnlySheet_(sh, invitedCount) {
  // Explanation banner in row 2 (unused on regular event tabs).
  sh.getRange(2, 1, 1, 26).breakApart();
  sh.getRange('A2:E2').merge()
    .setValue('Invite-only event: list every invited resident in column A and tick "Attended?" for those who came. ' +
      'Only ticked residents count toward Events Attended AND Eligible Events - unticked invitees and everyone else are unaffected.')
    .setFontStyle('italic').setFontColor('#666666').setWrap(true).setVerticalAlignment('top');
  sh.setRowHeight(2, 36);

  // Event Type marker - this is what the rest of the script checks (isInviteOnlySheet_).
  sh.getRange('A6:B6').setValues([['Event Type', INVITE_ONLY_TYPE]]);
  sh.getRange('A6').setBackground('#D9EAF7').setFontWeight('bold');
  sh.getRange('B6').setBackground('#FCE5CD').setFontWeight('bold');
  sh.getRange('B6').setNote('Do not change. This marks the tab as invite-only: only rows ticked in column B are counted.');

  // Live counts instead of typed numbers. D3 ("Venue") and D6 ("Applies To
  // Dorm(s)") keep their labels so ensureEventSheetMetaLabels_ leaves this tab alone.
  sh.getRange('D4:D5').setValues([['Invited'], ['Attended (ticked)']]);
  sh.getRange('E4').setFormula('=COUNTA(A' + ATTENDEE_START_ROW + ':A)');
  sh.getRange('E5').setFormula('=COUNTIF(B' + ATTENDEE_START_ROW + ':B,TRUE)');
  sh.getRange('E4:E5').setBackground('#EFEFEF').setNumberFormat('0');
  sh.getRange('E6').setValue('Invited residents only').setFontColor('#666666').setFontStyle('italic');
  sh.getRange('E6').setNote('Ignored for invite-only events - the invite list in column A is the scope.');

  sh.getRange('B7').setValue('Attended?');
  sh.getRange('E7').setValue('Notes');
  sh.getRange('B' + ATTENDEE_START_ROW + ':B1000').setBackground('#FCE5CD');
  applyInviteCheckboxes_(sh, ATTENDEE_START_ROW, invitedCount + INVITE_CHECKBOX_BUFFER_ROWS);
}

/** Adds the "Attended?" checkbox rule to column B for the given rows. Rule only - existing ticks are never changed. */
function applyInviteCheckboxes_(sh, startRow, numRows) {
  if (!numRows || numRows < 1) return;
  const rule = SpreadsheetApp.newDataValidation().requireCheckbox().build();
  sh.getRange(startRow, INVITE_ATTENDED_COL, numRows, 1).setDataValidation(rule).setHorizontalAlignment('center');
}

/** True if this event tab is invite-only (B6 reads "Invite-Only"). Regular tabs have A6:B6 blank. */
function isInviteOnlySheet_(sheet) {
  return String(sheet.getRange('B6').getValue()).trim().toLowerCase() === INVITE_ONLY_TYPE.toLowerCase();
}

/** Reads a checkbox / pasted tick. TRUE from a real checkbox, or text like TRUE / Yes / Y / x / 1 / ✓ / Attended. */
function isTicked_(v) {
  if (v === true) return true;
  if (v === false || v === null || v === undefined) return false;
  return /^(true|yes|y|x|1|\u2713|\u2714|\u2611|attended|present|came)$/i.test(String(v).trim());
}

/** Rewrites the on-sheet workflow / quick-reference tables on the Dashboard so they match the current menu. Safe to re-run. */
function refreshDashboardInstructions() {
  const sh = ensureSheet_(DASHBOARD_SHEET);

  sh.getRange('A7:B13').clearContent();
  sh.getRange('D7:F13').clearContent();

  sh.getRange('A7').setValue('Student Staff Workflow');
  sh.getRange('A8:B13').setValues([
    [1, 'Open this shared Google Sheet.'],
    [2, 'Dorm Attendance -> Events -> Create New Event Sheet (or Import Event to bulk-import a list; Import Invite-Only Event for small invite-only events, then tick who came).'],
    [3, 'Enter the event name and event date when prompted.'],
    [4, 'Enter attendee student ID numbers in Column A of the generated event tab, starting at row 8.'],
    [5, 'Dorm Attendance -> Events -> Refresh Attendance Log.'],
    [6, 'Review dorm summaries on the relevant dorm tab.']
  ]);

  sh.getRange('D7').setValue('Quick Command Reference');
  sh.getRange('D8:F13').setValues([
    ['Register a new student', 'Students -> Register Student (Single)', 'Adds one resident to the registry.'],
    ['Register many students at once', 'Students -> Bulk Register (Paste from Excel)', 'Paste ID / Dorm / Start Date / Notes rows.'],
    ['A student moves to a different dorm', 'Students -> Transfer Student (Change Dorm)', 'Keeps their attendance history - does not reset it.'],
    ['A student moves out / leaves', 'Students -> Deactivate Student (or Bulk Deactivate)', 'Keeps history, hides them from active rosters.'],
    ['Check data for mistakes', 'Maintenance -> Run Data Audit', 'Lists duplicate IDs, unknown dorms, bad dates, etc.'],
    ['Start a new academic year', 'Events -> Archive Previous Years\u2019 Events', 'Moves old logs out of the way and frees up Event IDs.']
  ]);

  styleHeader_(sh.getRange('A7:B7'));
  styleHeader_(sh.getRange('D7:F7'));
  sh.autoResizeColumns(1, 6);
}

/**
 * Rebuilds the "Current Summary" block on the Dashboard (rows 15-17, plus the
 * per-dorm table from row 20 down) - one row per currently ACTIVE dorm, each
 * pulling live from that dorm's own summary sheet. This section predates the
 * Apps Script project (it was hand-built directly in the Sheet before any of
 * this code existed) and nothing used to keep it in sync, which is why Add
 * Dorm never extended it: there was no code path that touched it at all.
 *
 * Deactivated dorms are dropped from the per-dorm table (their own sheet and
 * data are untouched, they just stop appearing in this particular rollup) -
 * consistent with how deactivated dorms already disappear from registration
 * dropdowns elsewhere. Called from Add Dorm, Deactivate Dorm, and Setup so
 * the table never goes stale.
 */
function refreshDashboardSummary_() {
  const sh = ensureSheet_(DASHBOARD_SHEET);
  const activeDorms = getActiveDormNames_();

  // breakApart() first: if A15:H15 is already merged in a shape that doesn't
  // exactly match (or something in that row got merged differently at some
  // point), calling merge() directly throws "You must select all cells in a
  // merged range..." and aborts the whole calling function (Add Dorm, etc.)
  // partway through. breakApart() on an ordinary unmerged range is a no-op,
  // so this is safe to run every time.
  //
  // BUGFIX (v6.5): breakApart() itself has the same "must select all cells"
  // restriction as merge() - if an existing merge only partially overlaps
  // the range you call breakApart() on (e.g. the real merge is wider than
  // A15:H15), breakApart() throws the exact same error instead of fixing
  // anything. Widened to the full row so it's guaranteed to fully contain
  // whatever merge shape actually exists, however wide that turns out to be.
  //
  // BUGFIX (v6.6): breakApart() only undoes the merge - any VALUE that was
  // sitting in a cell hidden underneath that merge (e.g. a second "Eligible
  // Event Slots" label someone typed into F20 years ago, invisible the whole
  // time it lived under a wider merged header) becomes visible again once
  // unmerged, without me ever having written it. That's what produced the
  // duplicate column and the stray percent-formatted value in the last
  // screenshot. Now does a full clearContent()+clearFormat() over the whole
  // region right after breakApart(), so anything revealed gets wiped before
  // this function rebuilds it cleanly from scratch.
  sh.getRange(15, 1, 3, 26).breakApart();
  sh.getRange(15, 1, 3, 26).clearContent().clearFormat();

  const headerRange = sh.getRange('A15:H15');
  headerRange.merge().setValue('Current Summary');
  headerRange.setBackground('#0B1F33').setFontColor('#FFFFFF').setFontWeight('bold').setHorizontalAlignment('center');

  sh.getRange('A16').setValue('Current Academic Year Start');
  sh.getRange('B16').setFormula('=Settings!B3').setNumberFormat('yyyy-mm-dd');
  sh.getRange('D16').setValue('Total Events This Year');
  sh.getRange('E16').setFormula('=COUNTIF(EventLog!C:C,">="&Settings!B3)').setNumberFormat('0');
  sh.getRange('G16').setValue('Total Attendance Records');
  sh.getRange('H16').setFormula('=MAX(0,COUNTA(AttendanceLog!A:A)-1)').setNumberFormat('0');

  sh.getRange('A17').setValue('Total Active Students');
  sh.getRange('B17').setFormula('=COUNTIF(StudentRegistry!D:D,"Active")').setNumberFormat('0');
  sh.getRange('D17').setValue('Dorms Tracked');
  // Was a typed-in static number before (never updated itself); now a live
  // formula against the Settings dorm table so it can't go stale again.
  sh.getRange('E17').setFormula('=COUNTIF(Settings!B' + DORM_TABLE_START_ROW + ':B1000,"Active")').setNumberFormat('0');
  sh.getRange('G17').setValue('Last Refreshed');
  sh.getRange('H17').setFormula('=Settings!B4').setNumberFormat('yyyy-mm-dd hh:mm');

  // Same breakApart() + full clear precaution for the per-dorm table area -
  // see the longer comment above for why both steps are necessary.
  sh.getRange(20, 1, 41, 26).breakApart();
  sh.getRange(20, 1, 41, 26).clearContent().clearFormat();

  // CHANGED (v6.8): blank spacer column at E, "Eligible Event Slots" moved
  // to F - matches the label / value / blank-spacer / label / value pattern
  // already used two rows up (D16/E16, blank F16, G16/H16).
  sh.getRange('A20:F20').setValues([['Dorm', 'Active Students', 'Group Attendance Score', 'Events Attended', '', 'Eligible Event Slots']]);
  styleHeader_(sh.getRange('A20:F20'));

  if (activeDorms.length) {
    const rows = activeDorms.map(dorm => {
      const ref = "'" + dorm.replace(/'/g, "''") + "'";
      return [
        dorm,
        '=COUNTIF(' + ref + '!B5:B1000,"Active")',
        '=' + ref + '!B2',
        '=SUM(' + ref + '!E5:E1000)',
        '',
        '=SUM(' + ref + '!F5:F1000)'
      ];
    });
    sh.getRange(21, 1, rows.length, 6).setValues(rows);
    // BUGFIX (v6.7): Active Students / Events Attended / Eligible Event Slots
    // are plain counts, but after clearFormat() they were picking up a "0.00%"
    // format from somewhere (most likely inherited from this same table's own
    // pre-existing state, predating this project) - so a raw count like 4
    // displayed as "400.00%". clearFormat() alone isn't enough to guarantee a
    // format; every numeric column here now gets its number format set
    // explicitly and unconditionally, so nothing can be inherited again.
    sh.getRange(21, 2, rows.length, 1).setNumberFormat('0');
    sh.getRange(21, 3, rows.length, 1).setNumberFormat('0.0%');
    sh.getRange(21, 4, rows.length, 1).setNumberFormat('0');
    sh.getRange(21, 6, rows.length, 1).setNumberFormat('0');
  }
}

function setupSettings_() {
  const sh = ensureSheet_(SETTINGS_SHEET);

  // Preserve any existing dorm configuration before repairing the sheet.
  const existingDormRows = readDormSettings_(sh, true);
  const existingSettings = {};
  if (sh.getLastRow() >= 2) {
    const vals = sh.getRange(1, 1, Math.min(sh.getLastRow(), 7), 2).getValues();
    vals.forEach(r => {
      if (r[0]) existingSettings[String(r[0]).trim()] = r[1];
    });
  }

  sh.clear();
  sh.getRange('A1:B1').setValues([['Setting', 'Value']]);
  styleHeader_(sh.getRange('A1:B1'));
  sh.getRange('A2:B7').setValues([
    ['Annual Reset Month', existingSettings['Annual Reset Month'] || 9],
    ['Current Academic Year Start', ''],
    ['Last Refreshed', existingSettings['Last Refreshed'] || ''],
    ['Attendance ID Entry Starts Row', ATTENDEE_START_ROW],
    ['Attendance ID Entry Column', ATTENDEE_COL],
    ['Template Sheet Name', TEMPLATE_SHEET]
  ]);
  sh.getRange('B3').setFormula('=DATE(IF(MONTH(TODAY())>=B2,YEAR(TODAY()),YEAR(TODAY())-1),B2,1)');
  sh.getRange('B3:B4').setNumberFormat('yyyy-mm-dd hh:mm');

  sh.getRange(DORM_TABLE_START_ROW - 2, 1).setValue('Dorm Configuration').setFontWeight('bold');
  sh.getRange(DORM_TABLE_START_ROW - 1, 1, 1, DORM_TABLE_HEADERS.length).setValues([DORM_TABLE_HEADERS]);
  styleHeader_(sh.getRange(DORM_TABLE_START_ROW - 1, 1, 1, DORM_TABLE_HEADERS.length));

  const dormRowsByName = {};
  existingDormRows.forEach(r => dormRowsByName[r.name] = r);
  DEFAULT_DORMS.forEach(name => {
    if (!dormRowsByName[name]) {
      dormRowsByName[name] = { name: name, status: 'Active', createdAt: new Date(), deactivatedAt: '', notes: 'Default dorm' };
    }
  });

  const outputRows = Object.keys(dormRowsByName)
    .sort((a, b) => a.localeCompare(b))
    .map(name => {
      const r = dormRowsByName[name];
      return [r.name, r.status || 'Active', r.createdAt || new Date(), r.deactivatedAt || '', r.notes || ''];
    });

  if (outputRows.length) {
    sh.getRange(DORM_TABLE_START_ROW, 1, outputRows.length, DORM_TABLE_HEADERS.length).setValues(outputRows);
    sh.getRange(DORM_TABLE_START_ROW, 2, Math.max(outputRows.length, 1), 1)
      .setDataValidation(SpreadsheetApp.newDataValidation().requireValueInList(['Active', 'Inactive'], true).build());
    sh.getRange(DORM_TABLE_START_ROW, 3, outputRows.length, 2).setNumberFormat('yyyy-mm-dd hh:mm');
  }

  sh.setFrozenRows(1);
  sh.autoResizeColumns(1, 5);
}

function setupRegistry_() {
  const sh = ensureSheet_(REGISTRY_SHEET);
  if (sh.getLastRow() === 0 || sh.getRange('A1').getValue() !== 'Student ID') {
    sh.clear();
    sh.getRange('A1:F1').setValues([['Student ID', 'Student Name / Notes', 'Dorm', 'Status', 'Start Date', 'Last Updated']]);
  }
  styleHeader_(sh.getRange('A1:F1'));
  sh.getRange('C2:C').setDataValidation(SpreadsheetApp.newDataValidation().requireValueInList(getActiveDormNames_(), true).build());
  sh.getRange('D2:D').setDataValidation(SpreadsheetApp.newDataValidation().requireValueInList(['Active', 'Inactive'], true).build());
  sh.getRange('E2:F').setNumberFormat('yyyy-mm-dd hh:mm');

  // Per-student attendance stats (covers every student - active or inactive,
  // any dorm - so this is the one place to see the whole roster at once).
  sh.getRange('G1:J1').setValues([['Events Attended', 'Eligible Events', 'Attendance %', 'Last Attended']]);
  styleHeader_(sh.getRange('G1:J1'));
  sh.getRange(1, 7, Math.max(sh.getMaxRows(), 1), 4).setBorder(null, true, null, null, null, null); // Left border - visually separates the stats section from the core registry columns.
  applyRegistryStatsFormulas_();

  sh.setFrozenRows(1);
  sh.autoResizeColumns(1, 10);
}

function setupEventLog_() {
  const sh = ensureSheet_(EVENT_LOG_SHEET);
  // BUGFIX (v4): this used to call sh.clear() unconditionally, which wiped
  // every logged event any time Setup / Repair Workbook was re-run. It now
  // only resets the sheet when the header is missing or wrong, matching the
  // safer pattern already used by setupRegistry_().
  if (sh.getLastRow() === 0 || sh.getRange('A1').getValue() !== 'Event ID') {
    sh.clear();
  }
  // Migration from v5 (no "Applies To") and v6.x (no "Event Type"): (re)write
  // the header row only - existing rows keep their data, and a blank Event
  // Type simply means "regular event", which is correct for everything
  // logged before v7.0.
  sh.getRange(1, 1, 1, EVENT_LOG_COLS).setValues([EVENT_LOG_HEADERS]);
  styleHeader_(sh.getRange(1, 1, 1, EVENT_LOG_COLS));
  sh.getRange('C:C').setNumberFormat('yyyy-mm-dd');
  sh.getRange('E:E').setNumberFormat('yyyy-mm-dd hh:mm');
  sh.setFrozenRows(1);
}

function setupAttendanceLog_() {
  const sh = ensureSheet_(ATTENDANCE_LOG_SHEET);
  // BUGFIX (v4): same unconditional sh.clear() issue as setupEventLog_ above -
  // re-running Setup used to silently delete the entire attendance history.
  if (sh.getLastRow() === 0 || sh.getRange('A1').getValue() !== 'Record ID') {
    sh.clear();
  }
  // Migration from v5/v6.x: header row only, existing rows kept as-is.
  sh.getRange(1, 1, 1, ATTENDANCE_LOG_COLS).setValues([ATTENDANCE_LOG_HEADERS]);
  styleHeader_(sh.getRange(1, 1, 1, ATTENDANCE_LOG_COLS));
  sh.getRange('C:C').setNumberFormat('yyyy-mm-dd');
  sh.getRange('H:H').setNumberFormat('yyyy-mm-dd hh:mm');
  sh.setFrozenRows(1);
}

function setupEventTemplate_() {
  const sh = ensureSheet_(TEMPLATE_SHEET);
  sh.clear();
  sh.getRange('A1:E1').merge().setValue('Event Attendance Sheet');
  sh.getRange('A1:E1').setBackground('#0B1F33').setFontColor('#FFFFFF').setFontWeight('bold').setHorizontalAlignment('center');
  sh.getRange('A3:B5').setValues([
    ['Event ID', ''],
    ['Event Date', ''],
    ['Event Name', '']
  ]);
  // Venue / Expected Attendees / Reported Attendee Count / Applies To Dorm(s):
  // optional metadata. Venue/Expected/Reported are filled in automatically by
  // the structured event import; Applies To is auto-detected from the event
  // name (see detectDormScope_) for both creation paths. All are editable -
  // staff can always override what was guessed.
  sh.getRange('D3:E6').setValues([
    ['Venue', ''],
    ['Expected Attendees', ''],
    ['Reported Attendee Count', ''],
    ['Applies To Dorm(s)', '']
  ]);
  sh.getRange('A3:A5').setBackground('#D9EAF7').setFontWeight('bold');
  sh.getRange('D3:D6').setBackground('#D9EAF7').setFontWeight('bold');
  sh.getRange('B3:B5').setBackground('#FFF2CC');
  sh.getRange('E3:E6').setBackground('#FFF2CC');
  sh.getRange('B4').setNumberFormat('yyyy-mm-dd').setDataValidation(SpreadsheetApp.newDataValidation().requireDate().build());
  sh.getRange('E6').setNote('Leave blank (or type "All") for an event open to every dorm. Otherwise, comma-separate the dorm names it applies to, e.g. "Dorm A, Dorm C". Auto-filled from the event name when possible - edit any time.');
  sh.getRange('A7:E7').setValues([['Student ID', 'Timestamp / Notes', 'Dorm Found', 'Registered?', 'Do Not Edit']]);
  styleHeader_(sh.getRange('A7:E7'));
  sh.getRange('A8:A1000').setBackground('#FFF2CC');
  applyEventLookupFormulas_(sh);
  sh.setFrozenRows(7);
  sh.setColumnWidths(1, 5, 140);
}

function setupTransferLog_() {
  const sh = ensureSheet_(TRANSFER_LOG_SHEET);
  if (sh.getLastRow() === 0 || sh.getRange('A1').getValue() !== 'Student ID') {
    sh.clear();
    sh.getRange('A1:D1').setValues([['Student ID', 'From Dorm', 'To Dorm', 'Transferred At']]);
  }
  styleHeader_(sh.getRange('A1:D1'));
  sh.getRange('D2:D').setNumberFormat('yyyy-mm-dd hh:mm');
  sh.setFrozenRows(1);
  sh.autoResizeColumns(1, 4);
}

/**
 * Creates the BulkActionTemplates sheet the first time only - a visible,
 * unprotected scratch sheet with one section per bulk action (matching the
 * exact columns each dialog expects), so staff can type or paste student
 * data into clean, validated cells, then select and copy just the data rows
 * straight into the matching dialog box.
 *
 * Guarded like setupRegistry_/setupEventLog_: once the sheet already looks
 * set up, this does nothing, so re-running Setup / Repair Workbook never
 * wipes out whatever staff currently have typed into it. Use Maintenance ->
 * Reset Bulk Action Templates to force a rebuild on purpose.
 */
function setupBulkActionTemplates_() {
  const sh = ensureSheet_(BULK_TEMPLATES_SHEET);
  if (String(sh.getRange('A1').getValue()).trim() !== 'Bulk Action Templates') {
    buildBulkActionTemplatesContent_(sh);
    return;
  }
  // v7.0: an existing sheet (set up by an older version) gets the new
  // invite-only section appended at the bottom - nothing above it is touched.
  const lastRow = sh.getLastRow();
  const colA = lastRow ? sh.getRange(1, 1, lastRow, 1).getValues().flat().map(v => String(v).trim()) : [];
  if (colA.indexOf(INVITE_TEMPLATE_TITLE) !== -1) return;

  // getLastRow() only sees cells with content, so unused yellow input rows of
  // section 3 could sit below it. Start after the last row that has ANY
  // background colour too, so the new section never overlaps old input cells.
  const bgs = sh.getRange(1, 1, sh.getMaxRows(), 5).getBackgrounds();
  let lastStyledRow = 0;
  bgs.forEach((r, i) => {
    if (r.some(c => c && c.toLowerCase() !== '#ffffff')) lastStyledRow = i + 1;
  });
  const startRow = Math.max(lastRow, lastStyledRow) + 2;
  const rowsNeeded = startRow + 45;
  if (sh.getMaxRows() < rowsNeeded) sh.insertRowsAfter(sh.getMaxRows(), rowsNeeded - sh.getMaxRows());
  sh.getRange(startRow, 1, 40, 5).clearDataValidations();
  writeInviteOnlyTemplateSection_(sh, startRow);
}

function resetBulkActionTemplates() {
  const ui = SpreadsheetApp.getUi();
  const confirm = ui.alert(
    'Reset Bulk Action Templates',
    'This clears anything currently typed into the BulkActionTemplates sheet and restores the blank headers/examples. Continue?',
    ui.ButtonSet.YES_NO
  );
  if (confirm !== ui.Button.YES) return;
  buildBulkActionTemplatesContent_(ensureSheet_(BULK_TEMPLATES_SHEET));
  ui.alert('Bulk Action Templates sheet reset.');
}

function buildBulkActionTemplatesContent_(sh) {
  sh.clear();
  sh.clearFormats();

  const activeDorms = getActiveDormNames_();
  const dormValidation = SpreadsheetApp.newDataValidation().requireValueInList(activeDorms, true).build();
  const dateValidation = SpreadsheetApp.newDataValidation().requireDate().build();

  let row = 1;
  sh.getRange(row, 1, 1, 5).merge().setValue('Bulk Action Templates')
    .setBackground('#0B1F33').setFontColor('#FFFFFF').setFontWeight('bold').setHorizontalAlignment('center');
  row++;
  sh.getRange(row, 1, 1, 5).merge()
    .setValue('Fill in a section below, select just the data rows (not the header or the italic example row), copy (Ctrl/Cmd+C), then open the matching dialog from the Dorm Attendance menu and paste (Ctrl/Cmd+V) into the box.')
    .setFontStyle('italic').setFontColor('#666666').setWrap(true);
  sh.setRowHeight(row, 34);
  row += 2;

  // 1. Bulk Register Students
  row = writeTemplateSection_(sh, row, {
    title: '1. Bulk Register Students  \u2014  Students -> Bulk Register (Paste from Excel)',
    headers: ['Student ID', 'Dorm', 'Start Date', 'Notes'],
    example: ['2024001', activeDorms[0] || 'Dorm A', '2026-04-01', 'Transfer student'],
    blankRows: 25,
    columnValidations: { 2: dormValidation, 3: dateValidation }
  });
  row++;

  // 2. Bulk Deactivate Students
  row = writeTemplateSection_(sh, row, {
    title: '2. Bulk Deactivate Students  \u2014  Students -> Bulk Deactivate (Paste List)',
    headers: ['Student ID'],
    example: ['2024001'],
    blankRows: 20,
    columnValidations: {}
  });
  row++;

  // 3. Import Event - the event's own details, one row...
  row = writeTemplateSection_(sh, row, {
    title: '3. Import Event  \u2014  Events -> Import Event with Attendance Details',
    headers: ['Event Name', 'Event Date', 'Venue', 'Attendee Count', 'Expected Attendee Count'],
    example: ['Takoyaki Party', '2026-07-11', 'Dorm B 1F', 5, 10],
    blankRows: 1,
    columnValidations: { 2: dateValidation }
  });
  row++; // blank row required before the Attendee ID table, matching the real paste format

  // ...then its attendee list, same section (no separate title).
  row = writeTemplateSection_(sh, row, {
    title: null,
    headers: ['Attendee ID', 'Attendee Dorm'],
    example: ['2024001', activeDorms[0] || 'Dorm A'],
    blankRows: 25,
    columnValidations: { 2: dormValidation }
  });
  row++;

  // 4. Import Invite-Only Event (v7.0)
  writeInviteOnlyTemplateSection_(sh, row);

  sh.setColumnWidths(1, 5, 160);
  sh.setFrozenRows(2);
}

/**
 * v7.0: section 4 of BulkActionTemplates - the invite-only event template.
 * Event details row, blank spacer row, then the invite list with a checkbox
 * column. No italic example rows here on purpose: staff copy this whole
 * block INCLUDING the headers (the importer needs them), so an example row
 * would get pasted in as if it were real data.
 */
function writeInviteOnlyTemplateSection_(sh, startRow) {
  const activeDorms = getActiveDormNames_();
  const dormValidation = SpreadsheetApp.newDataValidation().requireValueInList(activeDorms, true).build();
  const dateValidation = SpreadsheetApp.newDataValidation().requireDate().build();
  const checkboxValidation = SpreadsheetApp.newDataValidation().requireCheckbox().build();

  let row = writeTemplateSection_(sh, startRow, {
    title: INVITE_TEMPLATE_TITLE,
    note: 'For small events where only specific residents were invited. List everyone invited, tick "Attended?" for those who came ' +
      '(or leave them all unticked and tick them on the event tab later). Copy from the "Event Name" header down to the last invitee - ' +
      'headers included - and paste into the dialog. Only ticked residents get +1 Events Attended and +1 Eligible Events; ' +
      'invited no-shows and everyone else are unaffected.',
    headers: ['Event Name', 'Event Date', 'Venue'],
    example: null,
    blankRows: 1,
    columnValidations: { 2: dateValidation }
  });
  row++; // blank spacer row before the invite list, same as section 3

  row = writeTemplateSection_(sh, row, {
    title: null,
    headers: ['Invited ID', 'Dorm', 'Attended?'],
    example: null,
    blankRows: 30,
    columnValidations: { 2: dormValidation, 3: checkboxValidation }
  });
  return row;
}

/** Writes one template section (optional title, header row, italic example row, then N blank editable rows) starting at startRow. Returns the next free row. */
function writeTemplateSection_(sh, startRow, opts) {
  let row = startRow;
  const numCols = opts.headers.length;

  if (opts.title) {
    sh.getRange(row, 1, 1, 5).merge().setValue(opts.title).setBackground('#D9EAF7').setFontWeight('bold');
    row++;
  }

  // v7.0: optional explanatory note under the title.
  if (opts.note) {
    sh.getRange(row, 1, 1, 5).merge().setValue(opts.note)
      .setFontStyle('italic').setFontColor('#666666').setWrap(true).setVerticalAlignment('top');
    sh.setRowHeight(row, 48);
    row++;
  }

  sh.getRange(row, 1, 1, numCols).setValues([opts.headers]);
  styleHeader_(sh.getRange(row, 1, 1, numCols));
  row++;

  // v7.0: example row is now optional (pass example: null to skip it).
  if (opts.example && opts.example.length) {
    sh.getRange(row, 1, 1, opts.example.length).setValues([opts.example]);
    sh.getRange(row, 1, 1, opts.example.length).setFontStyle('italic').setFontColor('#999999');
    row++;
  }

  if (opts.blankRows) {
    const dataStartRow = row;
    sh.getRange(dataStartRow, 1, opts.blankRows, numCols).setBackground('#FFF2CC');
    Object.keys(opts.columnValidations || {}).forEach(colStr => {
      const col = Number(colStr);
      sh.getRange(dataStartRow, col, opts.blankRows, 1).setDataValidation(opts.columnValidations[colStr]);
    });
    row += opts.blankRows;
  }

  return row;
}

function setupDormSheets_() {
  const activeDorms = getActiveDormNames_();
  activeDorms.forEach(dorm => {
    const sh = ensureSheet_(dorm);
    setupOneDormSheet_(sh, dorm);
  });
}

function setupOneDormSheet_(sh, dorm) {
  // breakApart() defensively before merging - this runs on every active dorm
  // sheet on every registry-changing action, so any merge-shape mismatch
  // here would hit the widest possible range of Dorm menu actions. See the
  // longer explanation on refreshDashboardSummary_.
  sh.getRange(1, 1, 1, 26).breakApart();
  sh.getRange('A1:H1').merge().setValue(dorm + ' Attendance Summary');
  sh.getRange('A1:H1').setBackground('#0B1F33').setFontColor('#FFFFFF').setFontWeight('bold').setHorizontalAlignment('center');
  sh.getRange('A2').setValue('Dorm Group Attendance Score');
  sh.getRange('B2').setFormula('=IFERROR(SUM(E5:E)/SUM(F5:F),0)').setNumberFormat('0.0%').setFontWeight('bold');
  sh.getRange('D2').setValue('Academic Year Start');
  sh.getRange('E2').setFormula('=Settings!B3').setNumberFormat('yyyy-mm-dd');
  sh.getRange('A4:H4').setValues([['Student ID', 'Status', 'Dorm', 'Start Date', 'Events Attended', 'Eligible Events', 'Attendance %', 'Last Attended']]);
  styleHeader_(sh.getRange('A4:H4'));
  sh.setFrozenRows(4);
  sh.setColumnWidths(1, 8, 130);
  sh.getRange('G:G').setNumberFormat('0.0%');
  sh.getRange('D:D').setNumberFormat('yyyy-mm-dd');
  sh.getRange('H:H').setNumberFormat('yyyy-mm-dd');
}

function applyEventSheetFormatting_(sh) {
  sh.getRange('A1:E1').setBackground('#0B1F33').setFontColor('#FFFFFF').setFontWeight('bold').setHorizontalAlignment('center');
  sh.getRange('A3:A5').setBackground('#D9EAF7').setFontWeight('bold');
  sh.getRange('D3:D6').setBackground('#D9EAF7').setFontWeight('bold');
  sh.getRange('B3:B5').setBackground('#FFF2CC');
  sh.getRange('E3:E6').setBackground('#FFF2CC');
  sh.getRange('B4').setNumberFormat('yyyy-mm-dd').setDataValidation(SpreadsheetApp.newDataValidation().requireDate().build());
  styleHeader_(sh.getRange('A7:E7'));
  sh.getRange('A8:A1000').setBackground('#FFF2CC');
  applyEventLookupFormulas_(sh);
  sh.setFrozenRows(7);
}

function applyEventLookupFormulas_(sh) {
  const startRow = ATTENDEE_START_ROW;
  const endRow = 1000;
  const numRows = endRow - startRow + 1;

  // Clear old ARRAYFORMULA spill ranges and any stale values/errors in backend lookup columns.
  sh.getRange(startRow, 3, numRows, 2).clearContent();

  const formulas = [];
  for (let r = startRow; r <= endRow; r++) {
    formulas.push([
      '=IF($A' + r + '=\"\",\"\",IFERROR(VLOOKUP($A' + r + ',StudentRegistry!$A:$D,3,FALSE),\"Not found\"))',
      '=IF($A' + r + '=\"\",\"\",IFERROR(IF(VLOOKUP($A' + r + ',StudentRegistry!$A:$D,4,FALSE)=\"Active\",\"Yes\",\"Inactive\"),\"No\"))'
    ]);
  }

  sh.getRange(startRow, 3, numRows, 2).setFormulas(formulas);
  sh.getRange(startRow, 3, numRows, 2).setBackground('#FFFFFF');
}

function writeEventLog_(eventRows) {
  const sh = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(EVENT_LOG_SHEET);
  sh.clearContents();
  sh.getRange(1, 1, 1, EVENT_LOG_COLS).setValues([EVENT_LOG_HEADERS]);
  if (eventRows.length) {
    // Remove duplicate event IDs, keeping the last encountered entry.
    const byId = {};
    eventRows.forEach(row => byId[row[0]] = row);
    const uniqueRows = Object.keys(byId).sort().map(id => byId[id]);
    sh.getRange(2, 1, uniqueRows.length, EVENT_LOG_COLS).setValues(uniqueRows);
  }
  styleHeader_(sh.getRange(1, 1, 1, EVENT_LOG_COLS));
}

function writeAttendanceLog_(attendanceRows) {
  const sh = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(ATTENDANCE_LOG_SHEET);
  sh.clearContents();
  sh.getRange(1, 1, 1, ATTENDANCE_LOG_COLS).setValues([ATTENDANCE_LOG_HEADERS]);
  if (attendanceRows.length) {
    attendanceRows.sort((a, b) => String(a[1]).localeCompare(String(b[1])) || String(a[4]).localeCompare(String(b[4])));
    sh.getRange(2, 1, attendanceRows.length, ATTENDANCE_LOG_COLS).setValues(attendanceRows);
  }
  styleHeader_(sh.getRange(1, 1, 1, ATTENDANCE_LOG_COLS));
}

function upsertEventLog_(eventId, eventName, eventDate, sheetName, dormScope, eventType) {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sh = ss.getSheetByName(EVENT_LOG_SHEET) || ss.insertSheet(EVENT_LOG_SHEET);
  if (sh.getLastRow() < 1) setupEventLog_();

  const lastRow = sh.getLastRow();
  const ids = lastRow > 1 ? sh.getRange(2, 1, lastRow - 1, 1).getValues().flat().map(String) : [];
  const idx = ids.indexOf(eventId);
  const rowValues = [[eventId, eventName, eventDate, sheetName, new Date(), dormScope || '', eventType || '']];

  if (idx >= 0) {
    sh.getRange(idx + 2, 1, 1, EVENT_LOG_COLS).setValues(rowValues);
  } else {
    sh.appendRow(rowValues[0]);
  }
}

function getDormSettings_() {
  const sh = ensureSheet_(SETTINGS_SHEET);
  return readDormSettings_(sh, false);
}

function readDormSettings_(sh, tolerateMissing) {
  if (!sh || sh.getLastRow() < DORM_TABLE_START_ROW) return [];

  const headerRow = DORM_TABLE_START_ROW - 1;
  const headers = sh.getRange(headerRow, 1, 1, DORM_TABLE_HEADERS.length).getValues()[0].map(String);
  const looksLikeDormTable = headers[0] === DORM_TABLE_HEADERS[0] && headers[1] === DORM_TABLE_HEADERS[1];
  if (!looksLikeDormTable) return [];

  const numRows = sh.getLastRow() - DORM_TABLE_START_ROW + 1;
  if (numRows <= 0) return [];

  const values = sh.getRange(DORM_TABLE_START_ROW, 1, numRows, DORM_TABLE_HEADERS.length).getValues();
  return values
    .filter(r => String(r[0]).trim())
    .map(r => ({
      name: String(r[0]).trim(),
      status: String(r[1] || 'Active').trim(),
      createdAt: r[2],
      deactivatedAt: r[3],
      notes: r[4]
    }));
}

function getActiveDormNames_() {
  let rows = getDormSettings_();
  if (!rows.length) rows = DEFAULT_DORMS.map(name => ({ name: name, status: 'Active' }));
  return rows
    .filter(r => String(r.status || 'Active').toLowerCase() === 'active')
    .map(r => r.name);
}

function getInactiveDormNames_() {
  return getDormSettings_()
    .filter(r => String(r.status || '').toLowerCase() === 'inactive')
    .map(r => r.name);
}

function upsertDormSetting_(dormName, status, notes) {
  const sh = ensureSheet_(SETTINGS_SHEET);
  setupSettingsIfNeeded_();
  const rows = getDormSettings_();
  const idx = rows.findIndex(r => r.name === dormName);
  const now = new Date();
  if (idx >= 0) {
    const row = DORM_TABLE_START_ROW + idx;
    const oldStatus = String(sh.getRange(row, 2).getValue()).trim();
    sh.getRange(row, 2).setValue(status);
    if (!sh.getRange(row, 3).getValue()) sh.getRange(row, 3).setValue(now);
    sh.getRange(row, 4).setValue(status === 'Inactive' ? now : '');
    if (notes) sh.getRange(row, 5).setValue(notes);
    return oldStatus === 'Inactive' && status === 'Active' ? 'reactivated' : 'updated';
  }
  sh.appendRow([dormName, status, now, status === 'Inactive' ? now : '', notes || '']);
  return 'added';
}

function markDormInactive_(dormName) {
  upsertDormSetting_(dormName, 'Inactive', 'Deactivated via menu; historical records preserved.');
}

function setupSettingsIfNeeded_() {
  const sh = ensureSheet_(SETTINGS_SHEET);
  const header = sh.getRange(DORM_TABLE_START_ROW - 1, 1).getValue();
  if (header !== DORM_TABLE_HEADERS[0]) setupSettings_();
}

function markInactiveDormSheet_(dormName) {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sh = ss.getSheetByName(dormName);
  if (!sh) return;
  sh.getRange(1, 1, 1, 26).breakApart();
  sh.getRange('A1:H1').merge().setValue(dormName + ' Attendance Summary (Inactive Dorm)');
  sh.getRange('A1:H1').setBackground('#666666').setFontColor('#FFFFFF').setFontWeight('bold').setHorizontalAlignment('center');
  sh.getRange('A3').setValue('This dorm is inactive. Its sheet and historical attendance records are preserved, but it is not available for new student registrations.');
  sh.getRange(3, 1, 1, 26).breakApart();
  sh.getRange('A3:H3').merge().setBackground('#F4CCCC').setFontWeight('bold');
}

function isReservedSheetName_(name) {
  const reserved = [
    DASHBOARD_SHEET, REGISTRY_SHEET, EVENT_LOG_SHEET, ATTENDANCE_LOG_SHEET, SETTINGS_SHEET, TEMPLATE_SHEET,
    TRANSFER_LOG_SHEET, DELETED_STUDENTS_SHEET, EVENT_LOG_ARCHIVE_SHEET, ATTENDANCE_LOG_ARCHIVE_SHEET,
    DATA_AUDIT_SHEET, BULK_TEMPLATES_SHEET, 'Event_Example'
  ];
  return reserved.indexOf(name) >= 0 || name.indexOf(EVENT_PREFIX) === 0;
}

function getRegistryMap_() {
  const rows = getRegistryRows_();
  const map = {};
  rows.forEach(r => map[r.studentId] = r);
  return map;
}

function getRegistryRows_() {
  const sh = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(REGISTRY_SHEET);
  if (!sh || sh.getLastRow() < 2) return [];
  const values = sh.getRange(2, 1, sh.getLastRow() - 1, 6).getValues();
  return values
    .filter(r => normalizeId_(r[0]))
    .map(r => ({
      studentId: normalizeId_(r[0]),
      notes: r[1],
      dorm: String(r[2]).trim(),
      status: String(r[3]).trim() || 'Active',
      startDate: r[4] instanceof Date ? r[4] : getAcademicYearStart_(),
      lastUpdated: r[5]
    }));
}

function findRegistryRow_(studentId) {
  const sh = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(REGISTRY_SHEET);
  if (!sh || sh.getLastRow() < 2) return null;
  const values = sh.getRange(2, 1, sh.getLastRow() - 1, 1).getValues().flat();
  for (let i = 0; i < values.length; i++) {
    if (normalizeId_(values[i]) === studentId) return i + 2;
  }
  return null;
}

function getAcademicYearStart_() {
  const settings = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(SETTINGS_SHEET);
  if (settings) {
    const v = settings.getRange('B3').getValue();
    if (v instanceof Date) return v;
  }
  const now = new Date();
  const year = now.getMonth() + 1 >= 9 ? now.getFullYear() : now.getFullYear() - 1;
  return new Date(year, 8, 1);
}

function setLastRefreshed_(dateValue) {
  const sh = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(SETTINGS_SHEET);
  if (sh) sh.getRange('B4').setValue(dateValue).setNumberFormat('yyyy-mm-dd hh:mm');
}

function isEventSheet_(sheet) {
  const name = sheet.getName();
  if (name === TEMPLATE_SHEET) return false;
  if (name.indexOf(EVENT_PREFIX) === 0) return true;
  // Also accept event tabs if they have the event-sheet title and a valid event ID in B3.
  const title = String(sheet.getRange('A1').getValue()).trim();
  const eventId = String(sheet.getRange('B3').getValue()).trim();
  return title.indexOf('Event Attendance Sheet') >= 0 && eventId.indexOf(EVENT_PREFIX) === 0;
}

function ensureSheet_(name) {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  return ss.getSheetByName(name) || ss.insertSheet(name);
}

function styleHeader_(range) {
  range.setBackground('#1F4E78').setFontColor('#FFFFFF').setFontWeight('bold').setHorizontalAlignment('center');
}

function normalizeId_(value) {
  return String(value || '').trim().toUpperCase();
}

function parseDate_(text) {
  if (!text) return null;
  text = String(text).trim();
  // Accepts yyyy-mm-dd (e.g. 2026-07-11) and yyyy/m/d (e.g. 2026/6/26 - what
  // Sheets/Excel produces when you copy a date-formatted cell).
  const m = /^(\d{4})[-\/](\d{1,2})[-\/](\d{1,2})$/.exec(text);
  if (m) {
    const d = new Date(Number(m[1]), Number(m[2]) - 1, Number(m[3]));
    if (!isNaN(d.getTime())) return d;
  }
  const d = new Date(text);
  return isNaN(d.getTime()) ? null : d;
}

/**
 * Guesses which dorm(s) an event applies to from its name, e.g.:
 *   "Dorm D CCM"                  -> "Dorm D"
 *   "Dorm A/C/K CCM & Movie Night" -> "Dorm A,Dorm C,Dorm K"
 *   "Takoyaki Party"              -> ""  (no dorm mentioned - applies to everyone)
 * Returns a comma-separated (no spaces) canonical list, or "" for "all dorms".
 * Dorm codes that aren't currently configured (e.g. "Dorm K" before a "Dorm K"
 * dorm exists in Settings) are still included, not silently dropped - Data
 * Audit will flag them so staff notice and can add the dorm or fix a typo.
 */
function detectDormScope_(eventName) {
  const name = String(eventName || '');
  const allDorms = getActiveDormNames_().concat(getInactiveDormNames_());
  const matched = {};

  // Exact dorm-name matches - handles "Meikokan", "Dorm D", etc. as whole words.
  allDorms.forEach(dorm => {
    const pattern = new RegExp('(^|[^a-zA-Z0-9])' + escapeRegex_(dorm) + '($|[^a-zA-Z0-9])', 'i');
    if (pattern.test(name)) matched[dorm] = true;
  });

  // Compressed "Dorm A/C/K" or "Dorm A, C, K" style lists.
  const listPattern = /\bdorm\s+([a-zA-Z0-9]+(?:\s*[\/,]\s*[a-zA-Z0-9]+)+)/gi;
  let m;
  while ((m = listPattern.exec(name)) !== null) {
    m[1].split(/[\/,]/).map(t => t.trim()).filter(Boolean).forEach(t => {
      matched['Dorm ' + t] = true;
    });
  }

  return Object.keys(matched).sort().join(',');
}

function escapeRegex_(s) {
  return String(s).replace(/[.*+?^${}()|[\]\\]/g, '\\$&');
}

/** Cleans up a staff-typed or auto-detected dorm scope into the canonical comma-separated (no spaces) form. "All" / "All Dorms" / blank all mean "applies to everyone". */
function normalizeDormScope_(text) {
  if (!text) return '';
  const trimmed = String(text).trim();
  if (!trimmed || /^all(\s+dorms?)?$/i.test(trimmed)) return '';
  return trimmed.split(',').map(s => s.trim()).filter(Boolean).join(',');
}

/**
 * Generates an Event ID that combines the event's date with a short
 * sequence number, e.g. "EVT_20260626_01".
 *
 * CHANGED (v5): v4 shortened the ID to just a 2-digit sequence (e.g.
 * "EVT_01") with no date. That made IDs hard to place at a glance, so the
 * date is back, with the sequence still capped at 2 digits. The sequence
 * scans the live EventLog for the highest "..._NN" suffix already in use
 * (older formats - both the very first long timestamp IDs and last week's
 * date-less "EVT_01" style - are ignored, so nothing needs renaming) and
 * returns the next one, zero-padded to 2 digits. Archiving old events
 * (Events -> Archive Previous Years' Events) clears old IDs out of EventLog,
 * which frees the 2-digit range back up.
 */
function makeEventId_(eventDate) {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sh = ss.getSheetByName(EVENT_LOG_SHEET);
  let maxNum = 0;
  if (sh && sh.getLastRow() > 1) {
    const ids = sh.getRange(2, 1, sh.getLastRow() - 1, 1).getValues().flat();
    ids.forEach(id => {
      const m = /^EVT_\d{8}_(\d{2})$/.exec(String(id).trim());
      if (m) maxNum = Math.max(maxNum, parseInt(m[1], 10));
    });
  }
  const next = maxNum + 1;
  if (next > MAX_EVENT_ID_NUMBER) {
    throw new Error(
      'Reached the maximum of ' + MAX_EVENT_ID_NUMBER + ' active event IDs for this period (the sequence number is capped at ' + MAX_EVENT_ID_NUMBER + '). ' +
      'Run Events -> Archive Previous Years\u2019 Events to free up IDs before creating a new event.'
    );
  }
  const datePart = Utilities.formatDate(eventDate, Session.getScriptTimeZone(), 'yyyyMMdd');
  return EVENT_PREFIX + datePart + '_' + String(next).padStart(2, '0');
}

function sanitizeSheetName_(name) {
  return String(name)
    .replace(/[\[\]\*\?\/\\:]/g, '')
    .replace(/\s+/g, '_')
    .replace(/_+/g, '_')
    .substring(0, 60) || 'Event';
}

function uniqueSheetName_(baseName) {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  let name = baseName.substring(0, 99);
  let i = 2;
  while (ss.getSheetByName(name)) {
    const suffix = '_' + i;
    name = baseName.substring(0, 99 - suffix.length) + suffix;
    i++;
  }
  return name;
}