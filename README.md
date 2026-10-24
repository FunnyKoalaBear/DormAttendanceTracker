# Dorm Attendance Tracker

Google Apps Script automation for managing dorm residents, events, and attendance in a Google Sheets workbook.

The project provides a dashboard, per-dorm summaries, student attendance statistics, event sheets, bulk copy-and-paste workflows, and maintenance tools for keeping the workbook consistent.

## Repository contents

- `All Dorm Attendance Tracker.xlsx` - workbook template/data source to import or convert to Google Sheets.
- `Code.gs` - the bound Apps Script implementation.
- `BulkRegister.html` - dialog for registering multiple students from tab-separated data.
- `BulkDeactivate.html` - dialog for deactivating multiple students from a list of IDs.
- `ImportEvent.html` - dialog for importing event details and attendee data.

## Setup

1. Open or import `All Dorm Attendance Tracker.xlsx` in Google Sheets.
2. In the spreadsheet, open **Extensions -> Apps Script**.
3. Replace the contents of `Code.gs` with the repository's [`Code.gs`](./Code.gs).
4. Add the three HTML files to the Apps Script project with the same filenames:
   - `BulkRegister.html`
   - `BulkDeactivate.html`
   - `ImportEvent.html`
5. Save the Apps Script project and reload the spreadsheet.
6. Open **Dorm Attendance -> Maintenance -> Setup / Repair Workbook** once.
7. Authorize the script when Google prompts you.

The setup routine creates or repairs the dashboard, student registry, event and attendance logs, event template, dorm sheets, settings, and bulk-action templates. It also backfills supported event metadata and hides system sheets.

> The script is intended to be bound to the attendance workbook. It uses `@OnlyCurrentDoc` and does not need access to unrelated Google Sheets.

## Main workflows

### Students

Use **Dorm Attendance -> Students** to:

- Register one student or bulk-register students.
- Transfer a student to another dorm while retaining attendance history.
- Deactivate a student without removing their registry row or attendance history.
- Permanently delete a mistaken or duplicate student entry. This requires confirmation and archives related attendance records.

For bulk registration, paste tab-separated rows in this order:

```text
Student ID    Dorm    Start Date    Notes
```

`Start Date` and `Notes` are optional. A new student with no start date defaults to today; an existing student's current start date is preserved when the field is blank.

For bulk deactivation, paste one student ID per line.

### Events and attendance

Use **Dorm Attendance -> Events** to:

- Create a new event sheet from the standard event template.
- Import an event and its attendance details from copied spreadsheet data.
- Refresh the attendance log.
- Repair event-sheet formulas.
- Archive previous years' events.
- Remove references to deleted event sheets.

When creating an event, provide its name, date, and applicable dorms. Dorm scope can be detected from the event name and overridden manually. Events with no dorm scope apply to all dorms.

The event import dialog accepts a table in this layout:

```text
Event Name    Event Date    Venue    Attendee Count    Expected Attendee Count

Attendee ID    Attendee Dorm
2025B638      Dorm B
2025B639      Dorm B
```

The attendee section begins at the row labeled `Attendee ID`. Blank rows are allowed. The importer can also scan less-structured pasted text when an event name is supplied separately.

### Dorms and maintenance

Use **Dorm Attendance -> Dorms** to add or deactivate dorms and rebuild dorm summary sheets.

Use **Dorm Attendance -> Maintenance** to run a data audit, protect system sheets, reset bulk-action templates, or rerun workbook setup. Student IDs can be entered into event sheets directly; the built-in `onEdit` trigger refreshes attendance-related calculations when attendee data changes.

## Workbook sheets

The setup process maintains these logical areas:

- `Dashboard` - overall counts, instructions, and per-dorm summary.
- `StudentRegistry` - student IDs, dorms, status, start dates, notes, and calculated attendance metrics.
- One sheet per active dorm - dorm-specific summaries.
- Event sheets - event metadata and attendee IDs.
- `EventLog` and `AttendanceLog` - normalized backend records used by formulas and reports.
- `Settings` - configured dorms and workbook settings.
- `EventTemplate` and `BulkActionTemplates` - reusable templates.
- Audit, transfer, deletion, and archive sheets - operational history and recovery records.

Avoid changing backend columns or formula cells manually. Run **Maintenance -> Protect System Sheets** after the staff list and workbook structure are finalized.

## Upgrading

To upgrade an existing workbook, replace `Code.gs`, save, reload the spreadsheet, and run **Maintenance -> Setup / Repair Workbook**. The setup routine is designed to migrate existing event and attendance logs in place.

For workbooks upgraded from version 6.8 or earlier, delete the old `handleAttendanceEdit` and `handleSpreadsheetChange` installable triggers from the Apps Script editor. They were replaced by the built-in `onEdit` trigger in version 6.9.

## Notes

- The project has no local package manager, build step, or automated test suite; it runs in the Google Apps Script editor.
- Date input should use `yyyy-mm-dd` unless an import dialog explicitly documents another format.
- Keep a copy of the workbook before running structural repairs, archiving, or permanent deletion operations.
