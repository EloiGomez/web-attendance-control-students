**English** | [Català](README.ca.md) | [Español](README.es.md)

# Attendance Control

A lightweight school attendance tracker built with **Google Apps Script**, created to replace a slow, unwieldy multi-tab Excel spreadsheet used by an entire primary school to log daily student absences.

> Built as an AI pair-programming project with [Claude](https://claude.com/claude-code): I drove the requirements, design decisions, and testing (including catching a caching bug, a silent-failure bug, and a performance issue under load), while Claude handled the implementation.

## The problem

The original workflow was a single Excel file with one enormous grid per class (one row per student, one column per school day of the year — roughly 240 rows × 390 columns per class), duplicated across dozens of tabs for the whole school. Teachers had to type single-letter codes into cells by hand, the file took a long time to open and edit, and there was no automatic way to see a student's absence percentage for the year.

## The solution

A small web app served directly from Google Apps Script, backed by a Google Sheet used purely as a data store (never opened directly by teachers):

- **Take attendance by class and day**: pick a class and a date, see every student in a checkbox grid (morning / afternoon / justified / late-arrival), and save the whole day at once.
- **Live absence summary**: automatic attendance percentage per student (weighted — a missed morning counts differently than a missed afternoon, matching the school's real timetable), split by justified/unjustified absence, with a progressive color-coded alert level, a weekday-pattern breakdown, a search box, and sortable columns.
- **Access control**: restricted to an explicit allow-list of staff emails, checked server-side on every request — not just an obscure link.
- **Admin tooling**: a one-click end-of-year "promote all classes" tool (with an editable class-mapping sheet and a repeater flag per student), plus synthetic test-data generators for load-testing with hundreds of students.
- **Performance**: the backend batches all spreadsheet reads, reuses a single `Spreadsheet` handle per execution instead of re-fetching it repeatedly, and caches the computed summary (gzip-compressed to fit Apps Script's cache size limit) so repeat views are near-instant until the underlying data actually changes.
- Automatic dark mode (`prefers-color-scheme`), a mobile-friendly layout with sticky headers/columns, and no external dependencies — plain HTML/CSS/JS served through `HtmlService`.

## Stack

- **Backend**: Google Apps Script (JavaScript, V8 runtime) — `SpreadsheetApp`, `HtmlService`, `CacheService`, `PropertiesService`.
- **Frontend**: vanilla HTML/CSS/JS, no frameworks or build step — communicates with the backend via `google.script.run`.
- **Data store**: a Google Sheet, used as a simple structured database rather than a UI.

## Status

Functional prototype, tested with synthetic data (hundreds of students, thousands of attendance records). Contains no real student data — the `Students` sheet ships with example names only.

## Setup & deployment

No local environment or build step is needed — everything runs inside Google's own infrastructure.

1. Create a new blank spreadsheet at [sheets.google.com](https://sheets.google.com).
2. Open **Extensions → Apps Script**. This creates an Apps Script project bound to that spreadsheet.
3. In the default `Code.gs` file, delete the placeholder content and paste in this repo's [`Code.gs`](Code.gs).
4. Add a new **HTML** file (the `+` next to "Files"), name it exactly `Index` (no extension), and paste in this repo's [`Index.html`](Index.html).
5. Save the project.
6. In the function dropdown at the top, select **`setup`** and click **Run**. The first run will ask you to authorize the script — accept it. This creates the `Config`, `Students`, `Teachers`, `Records` and `Promotion` sheets with example data and sane defaults.
7. (Optional) Run **`generateTestStudents`** and **`generateTestRecords`** to load a few hundred synthetic students and attendance records, useful for trying out the summary view and sorting/search without typing anything by hand.
8. Add the emails of anyone who should have access to the **`Teachers`** sheet — if that sheet is empty, the app lets anyone in, so this step matters before sharing the link.
9. Go to **Deploy → New deployment**, choose type **Web app**. Set:
   - **Execute as**: *User accessing the web app* — so each visitor's own Google identity is what the access-control check (and the "updated by" audit column) actually sees.
   - **Who has access**: *Anyone with a Google account* (or *Anyone within [domain]*, if deploying under a Google Workspace organization).
10. Click **Deploy**, then open the resulting `.../exec` URL. The first time each user opens it, Google will ask them to authorize the app for their own account — that's expected, since it runs as each visitor rather than as the developer.

To ship a later code change, edit the files, save, then **Deploy → Manage deployments → ✏️ → Version: New version → Deploy** — the same `.../exec` URL keeps working, it just starts serving the updated code.

## Files

- [`Code.gs`](Code.gs) — server-side logic (routing, access control, data access, business logic, admin tools).
- [`Index.html`](Index.html) — the single-page frontend.
