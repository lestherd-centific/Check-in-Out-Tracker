// Reference copy of the "CheckInOut LogEvent v2" Office Script, which lives in
// Excel against Check In&Out Tracker.xlsx and is run by the
// "CheckInOut - Log Event" Power Automate flow.
//
// Nothing executes this file. If you change the script, change it in Excel
// first and then update this copy — a stale copy pasted back into Power
// Automate later would drop whatever column was added since, and it would
// look like it worked.
//
// Each parameter below becomes its own ScriptParameters/<name> on the flow's
// "Run script" action, mapped one-to-one from the Manual trigger's Request
// body — NOT a single JSON payload string. Map by NAME, not by position: the
// trigger schema lists destination before condition while the table has them
// the other way round.
//
// Append-only by design: every checkout and every checkin is its own row,
// linked by a shared checkoutId. Nothing is ever overwritten, so concurrent
// writes from many moderators at once never collide (no conflict-detection
// needed, unlike the Devices/Accessories/Props sync in the main tracker).
//
// Column order below MUST match the CheckoutLog table exactly:
//   id, checkoutId, eventType, ts, modName, deviceId, tag, model, serial,
//   project, location, condition, notes, destination, partner
//
// "destination" and "partner" are written LAST, after notes — Excel Online
// doesn't reliably keep a column in place if you drag or cut-paste it into
// the middle of an existing Table (it can silently snap back to the end on
// the next refresh), so rather than fight that, the script's order matches
// where new columns actually land: appended at the end.
//
// destination is optional and covers wherever equipment is actually being
// taken this trip, which may not exist anywhere in the main tracker's own
// Projects/Locations (a client site, an external shoot) — project+location
// stay the equipment's origin.
//
// partner is the moderator this person is paired with for the shift, taken
// from the Credentials list at sign-in. Blank means working alone, which is
// allowed — the site warns but never blocks. Both names sit on the one row
// rather than the event being written twice, so the checkoutId pairing that
// drives check-in stays intact.
//
// Every field is always a plain string — "" when not specified, never
// omitted, so the columns stay aligned.

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
  notes: string,
  destination: string, // "" when not specified
  partner: string      // "" when working alone
) {
  const table = workbook.getTable("CheckoutLog");
  table.addRow(-1, [
    id, checkoutId, eventType, ts, modName, deviceId, tag, model, serial,
    project, location, condition, notes, destination, partner
  ]);
}
