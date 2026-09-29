# Use Case: Edit a Court Case

End-user guide for the DigiDiggie TNG Access app.
Prerequisites: forms generated (`BuildAllForms`) and tables linked or
materialized — see [ACCESS-APP-SETUP.md](ACCESS-APP-SETUP.md).

Data model in one line: a **court case** has many **entries**; each entry links
to many **persons**; a case has **0..1 ruling**; each ruling has many **person
outcomes**.

---

## UC 1 — Open the case you want to edit

**Option A — browse list:** open `frmCourtCaseList`, navigate to the case
(the *Has ruling* column shows Yes/No), click **Open...**. The main form opens
on that case only.

**Option B — main form directly:** open `frmCourtCase` and page through cases
with **< Previous** / **Next >** (or the record navigator at the bottom).

## UC 2 — Edit the case header

Edit the fields: **Source** (combo), **Reference #**, **Case Year**, **District
Court**, **Court Case Text**.

Saving works like any Access form: a record is saved automatically when you
navigate away. The **Save Court Case** button (enabled only while there are
unsaved changes) saves the case *and* the ruling subform without navigating —
use it before closing the form.

> **Deleting a case** (use sparingly): click the record selector (the small box
> in the form's top-left corner) and press **Del**, or use the Delete (✗)
> button in the navigation bar at the bottom. What happens next depends on the
> setup: with **materialized local tables** Access cascades the delete to the
> case's entries, persons-links, ruling and outcomes. With **tables linked
> directly to PostgreSQL** the delete is *rejected* as long as the case has
> entries or a ruling (the database refuses to orphan them) — there is no UI
> for deleting those parts first, so ask a maintainer.

## UC 3 — Edit court case entries

Tab **Court Case Entries** shows the entries of the current case in a grid.

- **Quick edit:** click a grid cell and edit Year / Season / Land Use /
  Original Placename inline.
- **Add an entry:** click **Add Court Case Entry** (visible once the case is
  saved). A blank entry is created and its detail popup opens immediately.
- **Open an entry's detail:** select a row in the grid, click
  **Open Court Case Entry** (or the *Detail...* button in the grid).

### In the entry detail popup (`frmCourtCaseEntryDetail`)

- Edit **Entry Year**, **Season**, **Land Use**, **Original Placename**,
  **Curated Text**.
- **Standardised placename:** click **Pick...** → the placename search opens
  (first 50 names shown; type + **Search** for `LIKE` matching on placename or
  parish) → select a row → **Select**. The ID and name are written back.
- **Add a person:** **Add Person** → person search opens → search, or
  **New Person** to create one (Full Name is composed automatically from
  Given/Patronymic/Surname) → **Select**. The person is linked to this entry.
- **Remove a person:** select the row in the *People in Entry* grid, click
  **Delete Person**, confirm.

Close the popup when done; changes are saved on close/navigation.

### Editing entries in bulk

`frmCourtCaseEntriesList` lists **all** entries across all cases — useful to
standardise placenames across cases: pick an entry row and use its **Pick**
button, or **Open Detail**.

## UC 4 — Create and edit the ruling

A case can have **at most one ruling**. If none exists, the **Create Ruling**
button is visible on the main form:

1. Click **Create Ruling** — a ruling is created with a default type and the
   form jumps to the **Ruling** tab. The button disappears (one per case).
2. Edit **Year**, **Type**, **Legal Source**, **Description**.

The Ruling tab is empty when the case has no ruling.

## UC 5 — Record person outcomes (inside the ruling)

In the Ruling tab, the *Person Outcomes* grid lists outcomes of this ruling.

- **Add outcome:** click **Add Personal Outcome**. A row is created with
  outcome type *Okänd* (default) and focus jumps to it.
  Then set:
  - **Person** — combo listing only persons already linked to entries of
    *this case*. If the person is missing, add them to an entry first (UC 3).
  - **Outcome Type**, **Description**.
- **Delete outcome:** select the row, click **Delete Personal Outcome**,
  confirm.

---

## Rules & gotchas

| Topic | Behavior |
|---|---|
| One ruling per case | *Create Ruling* hides once a ruling exists; there is no ruling-delete UI |
| Outcome persons | Must already be linked to a case entry (combo is filtered to them) |
| Full Name | Auto-composed from name parts on save — don't type it |
| Autosave | Navigating away saves the current record; *Save Court Case* = explicit save incl. ruling subform |
| Unsaved case | Buttons that need a saved record (Add Entry, Create Ruling) are disabled until first save |
| Placename search | Shows top 50 on open; `LIKE %term%` on placename/parish after Search |

## Related docs

- Spec / field list: [../src/vba_tng_app/MINIMAL_APP.md](../src/vba_tng_app/MINIMAL_APP.md)
- Form interconnections: [../src/vba_tng_app/ARCHITECTURE_OVERVIEW.md](../src/vba_tng_app/ARCHITECTURE_OVERVIEW.md)
- Setup: [ACCESS-APP-SETUP.md](ACCESS-APP-SETUP.md) · Linking: [LINKED-DATABASE.md](LINKED-DATABASE.md)
