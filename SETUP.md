# Check-In/Out — Power Automate Setup

This new site is intentionally independent from the main Hardware Tracker's Excel workbook. It only reads live equipment/project/location data from the main tracker (via its existing Load flow) and writes its own append-only log to a brand-new workbook: **CheckInOut Data.xlsx** (included in this folder — upload it to SharePoint wherever the main tracker's workbook lives, then wire the two flows below to it).

No "Save" flow is needed. Every action (check out, check in) is a single appended row — never an overwrite — so concurrent moderators never collide and there's no conflict-detection logic to build, unlike the main tracker's Devices/Accessories/Props sync.

## 1. Workbook: CheckInOut Data.xlsx

Two tables, already created in the attached file:

**CheckoutLog** — columns, in this exact order:
`id, checkoutId, eventType, ts, modName, deviceId, tag, model, serial, project, location, condition, notes`

- `eventType` is either `checkout` or `checkin`.
- `checkoutId` links a checkout row to its later checkin row (same value on both).
- `condition` is `Good`, `Fair`, or `Damaged`.
- Delete the example row before going live.

**Credentials** — columns: `name, login, role`. Add your ~60 moderators here (role can just be `Mod` for all of them, or split further later). This is a separate table from the main tracker's Credentials — different population, different site.

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
    "condition": { "type": "string" },
    "notes": { "type": "string" }
  },
  "required": ["id","checkoutId","eventType","ts","modName","deviceId","tag","model","serial","project","location","condition","notes"]
}
```

**Action** — Excel Online (Business) → Run script, against CheckInOut Data.xlsx, script = `CheckoutLog-LogEvent.OfficeScript.ts` (in this folder). Map each `ScriptParameters/<field>` to the matching field from the trigger's dynamic content — one-to-one with the JSON schema above.

**Action** — Response, 200, empty body is fine (the site doesn't need anything back).

Once built, copy the flow's HTTP POST URL and paste it into `LOG_EVENT_URL` near the top of the new site's `index.html`.

## 3. Flow: "CheckInOut - Load Data"

Reads both tables back so the site can show current mods, currently-checked-out items, and recent history.

**Trigger** — Manual, no input required.

**Action** — Excel Online (Business) → List rows present in a table → Credentials.

**Action** — Excel Online (Business) → List rows present in a table → CheckoutLog.

**Action** — Response, 200, body:

```json
{
  "credentials": <dynamic content: value from the Credentials "List rows" action>,
  "checkoutLog": <dynamic content: value from the CheckoutLog "List rows" action>
}
```

(Same shape as your main tracker's "HWTracker - Load Data" flow — just fewer tables.)

Once built, copy the flow's HTTP GET/POST URL and paste it into `LOAD_URL` near the top of the new site's `index.html`.

## 4. One more URL: the main tracker's existing Load flow

The new site also fetches live Devices/Projects/Locations straight from your **existing, already-live** `HWTracker - Load Data` flow — no changes needed there. That URL is already filled in as `MAIN_TRACKER_LOAD_URL` in the new site's `index.html` (copied from your current site). Only the device's `status`, `bucket` (project), and `location` fields are used, purely read-only — this new site never writes back to that workbook.

## 5. What's deliberately deferred

Per your call to focus on logging first: there's no "currently checked out" dashboard, overdue flagging, or a supervised tab inside the main Hardware Management site yet. The CheckoutLog table has everything needed to build that later (it's a complete, query-able event history) — just not built out in this pass. The new site does show a simple "currently out" list so mods can find items to check back in, but that's the minimum needed to make the check-in flow work, not a reporting view.
