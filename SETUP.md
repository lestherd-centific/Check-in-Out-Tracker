# Check-In/Out — Setup and Architecture

This site is intentionally independent from the main Hardware Tracker's Excel workbook. It **reads** live equipment, project and location data from the main tracker (via that tracker's existing Load flow) and **writes** nothing back to it. Its own records go to two separate workbooks of its own.

Every action — check out, check in, issue report — is a single appended row, never an overwrite. That is what lets many moderators act at the same time with no conflict-detection logic, unlike the main tracker's Devices/Accessories/Props sync.

## Contents

1. Workbooks
2. Flows
3. Office Scripts in this folder
4. The main tracker's Load flow
5. Pagination — required
6. How the features map onto the data
7. What's deliberately deferred

---

## 1. Workbooks

Two, in different SharePoint locations, for a reason.

### 1a. CheckInOut Data.xlsx — *admin only*

Lives alongside the main tracker's workbook. Moderators never open it; they only reach it through the site.

**CheckoutLog** — columns, in this exact order:

```
id, checkoutId, eventType, ts, modName, deviceId, tag, model, serial,
project, location, condition, notes, destination, partner
```

- `eventType` is `checkout` or `checkin`.
- `checkoutId` links a checkout row to its later checkin row — the same value appears on both. This pairing is what "currently checked out" is derived from.
- `project` / `location` is the equipment's **origin** — where it is normally held, per the main tracker.
- `condition` is `Good`, `Fair` or `Damaged`.
- `destination` is optional: wherever the equipment is actually being taken this trip, which may not exist anywhere in the main tracker's Projects/Locations (a client site, an external shoot). Blank string when unspecified.
- `partner` is the moderator this person is paired with for the shift, blank when working alone. Both names sit on one row; the event is never written twice.
- `ts` is currently the submitting browser's local clock. See the note in section 7.

**Credentials** — `name, login, role`. The moderator roster. Separate from the main tracker's Credentials table: different population, different site. `role` is read but not enforced anywhere.

**Destinations** — single column, header `name`. An admin-maintained list of known destinations that appears as a dropdown at check-out. Not exhaustive by design — the checkout screen always offers "Other, type your own". A moderator's typed entry is recorded on that event but does **not** join the dropdown; someone with Excel access has to add it here.

### 1b. Device Issue Report.xlsx — *moderators have direct access*

On the moderators' SharePoint site, deliberately a separate **file**. SharePoint permissions are per file, and a hidden sheet is not a security boundary — so anything moderators must open directly has to live apart from Credentials and CheckoutLog.

**DeviceIssues** — columns, in this exact order:

```
Report ID, Date/Time Reported, Reported By, Project, Location, Device SN,
Asset Tag, Device Type, Description of Issue, Session/Context, Status,
Resolution Notes, Device Ref, Partner
```

- The site writes every column **except** `Status` and `Resolution Notes`, and never edits a row after writing it. Those two belong to the Hardware Tracking Owner and the Field Equipment Audit Lead from that point on, which is what keeps a triage decision safe from an app that is still appending rows.
- `Status` is `Open`, `In Review`, `Resolved` or `Replaced`, with a dropdown enforcing it. **Open and In Review withhold the device from check-out.** Resolved and Replaced release it. An unrecognised or blank status is treated as unresolved — holding a device back is the safer mistake.
- `Device Ref` is the main tracker's internal device id. It is the join key that drives the withholding, and it is blank for items the tracker does not hold (a radar cage, a hub, a power bank) — those are recorded but withhold nothing.
- `Date/Time Reported` is Pacific with the zone labelled, e.g. `9/24/2026, 8:54:20 AM PDT`.
- The workbook has an Instructions sheet explaining all of this to whoever fills in Status.

> **Adding a column to either table:** always append at the far **right**. Both Office Scripts write positionally, so a column inserted in the middle silently shifts every value after it. Excel Online also does not reliably keep a dragged column in place inside a Table — it can snap back to the end on a later refresh.

---

## 2. Flows

Four, in two pairs. Each pair has one Load (read) and one Log (append).

| Flow | Purpose | Site constant |
|---|---|---|
| CheckInOut - Load Data | Reads Credentials, CheckoutLog, Destinations | `LOAD_URL` |
| CheckInOut - Log Event | Appends one check-out or check-in row | `LOG_EVENT_URL` |
| Issues - Load Data | Reads DeviceIssues | `ISSUES_LOAD_URL` |
| Issues - Log Issue | Appends one issue report | `ISSUES_LOG_URL` |

**The issue flows are deliberately separate from the check-in/out flows**, rather than extra actions on the existing ones. The CheckInOut Load flow carries the Credentials list, so any failure in it stops everyone signing in. The issue workbook is one moderators can edit directly — someone renaming the table or moving the file is a realistic event. Keeping them apart means a problem with the issue log breaks the issue log and nothing else.

### 2a. CheckInOut - Log Event

**Trigger** — Manual, "When an HTTP request is received". Request Body JSON Schema:

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
    "notes": { "type": "string" },
    "partner": { "type": "string" }
  },
  "required": ["id","checkoutId","eventType","ts","modName","deviceId","tag","model","serial","project","location","destination","condition","notes","partner"]
}
```

**Action** — Excel Online (Business) → Run script, against `Check In&Out Tracker.xlsx`, script `CheckInOut LogEvent v2`. Map each `ScriptParameters/<field>` to the trigger field of the **same name** — not by position. The schema lists `destination` before `condition`, while the table has them the other way round; mapping down the list in order swaps them.

**Action** — Response, 200, empty body.

### 2b. CheckInOut - Load Data

**Trigger** — Manual. The site calls it with POST.

**Actions** — three × Excel Online (Business) → List rows present in a table: Credentials, CheckoutLog, Destinations.

**Action** — Response, 200:

```json
{
  "credentials": <value, from the Credentials List rows action>,
  "checkoutLog": <value, from the CheckoutLog List rows action>,
  "destinations": <value, from the Destinations List rows action>
}
```

Pick **value** from the dynamic content panel, not *body* — `body` wraps the array in a layer the site does not expect.

### 2c. Issues - Log Issue

**Trigger** — Manual. Request Body JSON Schema:

```json
{
  "type": "object",
  "properties": {
    "id": { "type": "string" },
    "ts": { "type": "string" },
    "reportedBy": { "type": "string" },
    "project": { "type": "string" },
    "location": { "type": "string" },
    "serial": { "type": "string" },
    "tag": { "type": "string" },
    "deviceType": { "type": "string" },
    "description": { "type": "string" },
    "context": { "type": "string" },
    "status": { "type": "string" },
    "resolutionNotes": { "type": "string" },
    "deviceRef": { "type": "string" },
    "partner": { "type": "string" }
  },
  "required": ["id","ts","reportedBy","project","location","serial","tag","deviceType","description","context","status","resolutionNotes","deviceRef","partner"]
}
```

**Action** — Run script against `Device Issue Report.xlsx`, script `DeviceIssues-LogIssue`. The site always sends `status` as `"Open"` and `resolutionNotes` as `""`.

**Action** — Response, 200, empty body.

### 2d. Issues - Load Data

**Trigger** — Manual, called with POST.
**Action** — List rows present in a table → DeviceIssues.
**Action** — Response, 200: `{ "issues": <value, from the List rows action> }`

Excel returns each column under its **header name verbatim**, spaces and all — `"Report ID"`, `"Session/Context"` — plus `@odata.etag` and `ItemInternalId`, which the site ignores. Renaming a header in the workbook will silently empty the matching field in the site, with no error anywhere. The mapping lives in one labelled block (`ISSUE_FIELDS`) in `index.html`; change both together.

---

## 3. Office Scripts in this folder

`CheckoutLog-LogEvent.OfficeScript.ts` and `DeviceIssues-LogIssue.OfficeScript.ts` are **reference copies** of the scripts that live in Excel. Nothing executes them from here.

If you change a script, change it in Excel first (Automate → the script → Save), then update the copy here. A stale copy is a trap: pasting it back into Power Automate later would drop whatever column was added since, and it would look like it worked.

---

## 4. The main tracker's Load flow

The site fetches live Devices, Projects and Locations from the **existing** `HWTracker - Load Data` flow — no changes needed there. Its URL is filled in as `MAIN_TRACKER_LOAD_URL` in `index.html`.

Only `status`, `bucket` (project), `location` and `kit` are used from each device, plus `name` and `archived` from each project. Purely read-only: this site never writes to that workbook.

`kit` matters more than it looks — it drives the one-tap "select the whole bundle" chips on both the Check Out and Check In screens. Devices sharing a `kit` value are treated as travelling together.

---

## 5. Pagination — required

Every "List rows present in a table" action needs **Settings → Pagination → On**, with a Threshold well above the table's expected size (10,000 is fine).

Without it the action returns only the **first 256 rows**, silently. Because Excel returns rows oldest-first, that means the site stops seeing the *newest* events — recently checked-out gear disappears from Check In and reappears as available. There is no error; it simply goes quiet.

Pagination is per-action and only appears when an action is selected. It exists on **List rows present in a table**, not on **Run script**.

---

## 6. How the features map onto the data

- **"Currently checked out"** is derived entirely from CheckoutLog: a `checkout` row whose `checkoutId` has no matching `checkin` row. The main tracker's own `status` field means something different — where equipment is deployed, not who is holding it this afternoon — and the two are deliberately independent.
- **A device is withheld from check-out** when DeviceIssues holds a row for its `Device Ref` whose `Status` is not Resolved or Replaced. It is shown greyed with the reason rather than hidden, so a moderator holding the physical item can see why.
- **Marking something Damaged at check-in** requires a note and files an issue report automatically, so the withholding rule has something to act on without anyone remembering a second step.
- **If the issue log is unreachable**, nothing is withheld and check-out carries on. A flow outage must not stop a shift.
- **Writes retry** up to three times on a dropped connection or a 429/5xx, but never on other errors — a deliberate rejection would only be repeated, and a blind retry of a request that actually landed would duplicate the row.

---

## 7. What's deliberately deferred

- **Timestamps on CheckoutLog are the submitting browser's local clock.** Issue reports are Pacific with the zone labelled; the check-out log is not, so the two are not directly comparable. A fix is written and tested but not shipped.
- **No notifications.** Overdue equipment and new issue reports are visible in the Hardware Tracker, but nothing emails or messages anyone — somebody has to look. The Log Issue flow is the natural place to add one.
- **Roles are not enforced.** The `role` column is read and displayed; nothing restricts what anyone can do.
- **Sign-in is attribution, not security.** A login name identifies who did something. Anyone who views the page source can see the flow URLs.
- **Destinations are maintained by hand** — see section 1a.
- **The moderator timeline in the Hardware Tracker reads `modName` only.** Until that repo reads the new `partner` column, a paired moderator still shows as idle there.
