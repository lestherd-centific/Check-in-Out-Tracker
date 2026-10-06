// Reference copy of the "CheckInOut LogEvent v2" Office Script, which lives in
// Excel against Check In&Out Tracker.xlsx and is run by the
// "CheckInOut - Log Event" Power Automate flow.
//
// Nothing executes this file. If you change the script, change it in Excel
// first (Automate → the script → Save) and then update this copy.
//
// Each parameter below becomes its own ScriptParameters/<name> on the flow's
// "Run script" action, mapped one-to-one from the Manual trigger's Request
// body — NOT a single JSON payload string.
//
// Append-only by design: every checkout and every checkin is its own row,
// linked by a shared checkoutId. Nothing is ever overwritten, so concurrent
// writes from many moderators at once never collide.
//
// WRITES BY HEADER NAME, NOT BY POSITION. The previous version mapped values
// to columns by their order in the sheet, which made every column change a
// breakage: a column inserted in the middle silently shifted every value
// after it, and adding one at the end broke writes outright until the script
// was updated in lockstep. Matching on the header row removes that whole
// class of fault — column order no longer matters at all.
//
// The two failure modes are deliberately asymmetric:
//   - A column this script doesn't know about is written blank. Somebody
//     adding a column must never stop a moderator checking equipment out.
//   - A value with no column to go in throws, naming the column. That means
//     a renamed or mistyped header fails loudly on the first write instead of
//     quietly emptying a field on every row from then on.
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
// drives check-in stays intact. Only a Moderation shift carries a partner; an
// Uploader row always has it blank.
//
// sessionRole is what this person is doing this shift: "Moderation" (running
// a session, normally paired) or "Uploader" (uploading into One Data, which
// is a solo job by nature). Deliberately NOT the same as the `role` column in
// Credentials, which is the roster role (Mod/Admin) — two different meanings
// of "role" in one system is how somebody eventually reads the wrong one.
// It is recorded per row rather than per person: the same moderator checks
// gear out as Moderation in the morning and as Uploader in the afternoon, and
// both rows should say so.
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
  partner: string,     // "" when working alone
  sessionRole: string  // "Moderation" | "Uploader"
) {
  const table = workbook.getTable("CheckoutLog");

  const byHeader: { [key: string]: string } = {
    id: id,
    checkoutId: checkoutId,
    eventType: eventType,
    ts: ts,
    modName: modName,
    deviceId: deviceId,
    tag: tag,
    model: model,
    serial: serial,
    project: project,
    location: location,
    condition: condition,
    notes: notes,
    destination: destination,
    partner: partner,
    sessionRole: sessionRole
  };

  const headers = table.getHeaderRowRange().getValues()[0].map(h => String(h).trim());

  const missing = Object.keys(byHeader).filter(k => headers.indexOf(k) === -1);
  if (missing.length > 0) {
    throw new Error("CheckoutLog is missing these columns: " + missing.join(", ") +
      " — check the header row spelling.");
  }

  table.addRow(-1, headers.map(h => (h in byHeader) ? byHeader[h] : ""));
}
