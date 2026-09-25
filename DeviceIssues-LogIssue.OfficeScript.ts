// Reference copy of the "DeviceIssues-LogIssue" Office Script, which lives in
// Excel against Device Issue Report.xlsx and is run by the
// "Issues - Log Issue" Power Automate flow.
//
// Nothing executes this file. If you change the script, change it in Excel
// first and then update this copy.
//
// Same pattern as CheckoutLog-LogEvent: each parameter below becomes its own
// ScriptParameters/<name> in the flow, mapped one-to-one from the Manual
// trigger's Request body — NOT a single JSON payload string.
//
// Append-only. The site writes a report once and never edits it afterwards,
// so two moderators filing at the same moment cannot collide and no
// conflict-detection is needed.
//
// Column order below MUST match the DeviceIssues table's column order exactly:
//   Report ID, Date/Time Reported, Reported By, Project, Location, Device SN,
//   Asset Tag, Device Type, Description of Issue, Session/Context, Status,
//   Resolution Notes, Device Ref, Partner
//
// status is always written as "Open" by the site, and resolutionNotes as "".
// Those two columns belong to the Audit Lead from that point on — nothing in
// the site ever writes to them again, which is what keeps a triage decision
// safe from being overwritten by an app that is still appending rows.
//
// Every parameter is a plain string, "" when not applicable, never omitted, so
// the columns stay aligned. deviceRef is "" for an item that isn't in the
// Hardware Tracker at all (a radar cage, a hub, a power bank) — such a report
// is recorded but blocks nothing, because there is no device to block.

function main(
  workbook: ExcelScript.Workbook,
  id: string,
  ts: string,              // Pacific, e.g. "9/23/2026, 2:41:07 PM PDT"
  reportedBy: string,
  project: string,
  location: string,
  serial: string,
  tag: string,
  deviceType: string,
  description: string,
  context: string,         // "Mid-session" | "During check-in" | "During check-out" | "Other"
  status: string,          // always "Open" from the site
  resolutionNotes: string, // always "" from the site
  deviceRef: string,       // "" when the item isn't in the Hardware Tracker
  partner: string          // "" when working alone
) {
  const table = workbook.getTable("DeviceIssues");
  table.addRow(-1, [
    id, ts, reportedBy, project, location, serial, tag, deviceType,
    description, context, status, resolutionNotes, deviceRef, partner
  ]);
}
