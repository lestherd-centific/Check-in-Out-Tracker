// Office Script for the "CheckInOut - Log Event" Power Automate flow.
// Paste this into the "Run script" action (same pattern as your existing
// "HWTracker - Log Activity" flow: each field below becomes its own
// ScriptParameters/<name> in the flow, mapped from the Manual trigger's
// Request body — NOT a single JSON payload string).
//
// Append-only by design: every checkout and every checkin is its own row,
// linked by a shared checkoutId. Nothing is ever overwritten, so concurrent
// writes from many moderators at once never collide (no conflict-detection
// needed, unlike the Devices/Accessories/Props sync in the main tracker).
//
// Column order below MUST match the CheckoutLog table's column order
// exactly: id, checkoutId, eventType, ts, modName, deviceId, tag, model,
// serial, project, location, condition, notes.

function main(
  workbook: ExcelScript.Workbook,
  id: string,
  checkoutId: string,
  eventType: string,   // "checkout" | "checkin"
  ts: string,
  modName: string,
  deviceId: string,
  tag: string,
  model: string,
  serial: string,
  project: string,
  location: string,
  condition: string,   // "Good" | "Fair" | "Damaged"
  notes: string
) {
  const table = workbook.getTable("CheckoutLog");
  table.addRow(-1, [
    id, checkoutId, eventType, ts, modName, deviceId, tag, model, serial,
    project, location, condition, notes
  ]);
}
