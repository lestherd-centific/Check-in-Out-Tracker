# Check-In/Out — Power Automate Setup

This new site is intentionally independent from the main Hardware Tracker's Excel workbook. It only reads live equipment/project/location data from the main tracker (via its existing Load flow) and writes its own append-only log to a brand-new workbook: **CheckInOut Data.xlsx**, which lives in SharePoint alongside the main tracker's workbook.

No "Save" flow is needed. Every action (check out, check in) is a single appended row — never an overwrite — so concurrent moderators never collide and there's no conflict-detection logic to build, unlike the main tracker's Devices/Accessories/Props sync.

## 1. Workbook: CheckInOut Data.xlsx

Three tables:

**CheckoutLog** — columns, in this exact order:
`id, checkoutId, eventType, ts, modName, deviceId, tag, model, serial, project, location, condition, notes, destination`

- `eventType` is either `checkout` or `checkin`.
- `checkoutId` links a checkout row to its later checkin row (same value on both).
- `project` / `location` is the equipment's **origin** — where it's normally held, per the main tracker.
- `condition` is `Good`, `Fair`, or `Damaged`.
- `destination` (added later, at the very end) is optional — wherever the equipment is actually being taken this trip, which may not be anywhere in the main tracker's own Projects/Locations (a client site, an external shoot location, etc). Blank string when not specified.
- Delete the example row before going live.

> **If this table already exists live:** add `destination` as a new column at the very **end** (after `notes`) — the Office Script's `table.addRow(...)` array writes columns positionally in the order above, with `destination` last, specifically so it can just be appended without disturbing anything before it. Don't try to drag/cut-paste it into the middle of the table to "match" a different order — Excel Online doesn't reliably keep a reordered column in place inside a Table (it can silently snap back on the next refresh), so the column staying at the end is the stable, correct state, not a problem to fix. Leave it blank for existing rows.

**Credentials** — columns: `name, login, role`. Add your ~60 moderators here (role can just be `Mod` for all of them, or split further later). This is a separate table from the main tracker's Credentials — different population, different site.

**Destinations** (new) — a single-column table, header `name`. A simple, admin-maintained list of known destinations (e.g. "Client Site — Seattle Office") that shows up as a dropdown when checking equipment out. Not required to be exhaustive — the checkout screen always has an "Other, type your own" option too, so this list is just a convenience, not a hard constraint. Add whatever destinations you already know about; grow it over time as new ones come up (there's currently no in-app way to add to this list — a mod's "Other" entry is a one-off, only reused automatically for future dropdown suggestions if someone with Excel access adds it here).

## 2. Flow: "CheckInOut - Log Event"

Appends exactly one row per action. Mirrors your existing "HWTracker - Log Activity" flow.

**Trigger** — Manual, "When an HTTP request is received" / Instant cloud flow, Request body JSON schema:

```json
{
  "type": "object",
  "properties": {
    "id": { "type": "string" },
    "checkoutId": { "type": "string" },
    "eventType": { "type": "string" },
    "ts": { "type": "string" },
    "modName": { "type": "string" },
    "deviceId": { "type": "string" },
    "tag": { "type": "string" },
    "model": { "type": "string" },
    "serial": { "type": "string" },
    "project": { "type": "string" },
    "location": { "type": "string" },
    "destination": { "type": "string" },
    "condition": { "type": "string" },
    "notes": { "type": "string" }
  },
  "required": ["id","checkoutId","eventType","ts","modName","deviceId","tag","model","serial","project","location","destination","condition","notes"]
}
```

**Action** — Excel Online (Business) → Run script, against CheckInOut Data.xlsx, script = `CheckoutLog-LogEvent.OfficeScript.ts` (in this folder — already updated to accept `destination`). Map each `ScriptParameters/<field>` to the matching field from the trigger's dynamic content — one-to-one with the JSON schema above.

> **If this flow already exists live:** (1) add `destination` to the trigger's Request Body JSON schema exactly as shown above (edit the schema directly, or resave the trigger with a sample payload that includes it), (2) paste the updated Office Script over the old one in the "Run script" step, and (3) a new `ScriptParameters/destination` field will appear — map it to `triggerBody()?['destination']` (or the equivalent dynamic content picker entry) the same way the existing fields are mapped.

**Action** — Response, 200, empty body is fine (the site doesn't need anything back).

Once built, copy the flow's HTTP POST URL and paste it into `LOG_EVENT_URL` near the top of the new site's `index.html`.

## 3. Flow: "CheckInOut - Load Data"

Reads both tables back so the site can show current mods, currently-checked-out items, and recent history.

**Trigger** — Manual, no input required.

**Action** — Excel Online (Business) → List rows present in a table → Credentials.

**Action** — Excel Online (Business) → List rows present in a table → CheckoutLog.

**Action** — Excel Online (Business) → List rows present in a table → Destinations.

**Action** — Response, 200, body:

```json
{
  "credentials": <dynamic content: value from the Credentials "List rows" action>,
  "checkoutLog": <dynamic content: value from the CheckoutLog "List rows" action>,
  "destinations": <dynamic content: value from the Destinations "List rows" action>
}
```

(Same shape as your main tracker's "HWTracker - Load Data" flow — just fewer tables.)

> **If this flow already exists live:** add the "List rows present in a table → Destinations" action (same pattern as the other two), and add `"destinations": <...>` to the existing Response action's body. The site treats a missing `destinations` field as an empty list (the "Other, type your own" option still works either way), so this can be rolled out whenever convenient — it's not a breaking change to skip for now.

Once built, copy the flow's HTTP GET/POST URL and paste it into `LOAD_URL` near the top of the new site's `index.html`.

## 4. One more URL: the main tracker's existing Load flow

The new site also fetches live Devices/Projects/Locations straight from your **existing, already-live** `HWTracker - Load Data` flow — no changes needed there. That URL is already filled in as `MAIN_TRACKER_LOAD_URL` in the new site's `index.html` (copied from your current site). Only the device's `status`, `bucket` (project), and `location` fields are used, purely read-only — this new site never writes back to that workbook.

## 5. What's deliberately deferred

Per your call to focus on logging first: there's no "currently checked out" dashboard, overdue flagging, or a supervised tab inside the main Hardware Management site yet. The CheckoutLog table has everything needed to build that later (it's a complete, query-able event history) — just not built out in this pass. The new site does show a simple "currently out" list so mods can find items to check back in, but that's the minimum needed to make the check-in flow work, not a reporting view.
