# Test Evidence Inserter

Proof-of-concept Excel task-pane add-in. Organizes screenshots after tester manually executes and validates tests.

## Architecture

```text
src/config/evidenceModules.ts   Typed module catalog
src/config/layoutProfiles.ts    Image, workbook, and layout configuration
src/services/imageValidation.ts MIME, signature, size, and decode validation
src/services/layoutEngine.ts    Pure pairing and geometry engine
src/services/shapeNames.ts      Managed shape names and sequences
src/services/excelService.ts    Office.js workbook reads and writes
src/App.tsx                     One modular task-pane UI
index.html                      Screenshot Inserter Docs page
documentation/site.css         Documentation page styling
public/assets/docs/             Documentation screenshots
manifest.xml                    Stable XML add-in manifest
manifest.local.xml              Local HTTPS development manifest
docs/                           GitHub Pages production bundle
```

Task pane checks same-origin `version.json` at startup, when window regains focus, and every 15 minutes. New releases show an **Update now** notice. Reload stays user-controlled and is blocked while a screenshot is pending, preventing clipboard-image loss. No cookies, browser history, or Office authentication state are cleared.

User guide: `https://thrddqno.github.io/screenshot_inserter/`. The task-pane **Help** link opens the guide in a new browser tab.

Documentation source lives outside generated `docs/`. Update `index.html`, `documentation/site.css`, or sanitized images in `public/assets/docs/`, then run `npm run build`. Keep screenshot filenames stable, use a demonstration workbook, and remove names, paths, URLs, credentials, and private data before capture.

Every evidence module remains available in Evidence Type dropdown. Tester chooses module manually; no automatic suggestion or scenario classification runs.

## Requirements

- Node.js 22 or newer
- OpenSSL for local certificate creation
- Excel on Windows or Excel for web
- Host support for `ExcelApi 1.10`

## Install

```bash
npm install
npm run certs:create
```

Trust `.certs/localhost.crt` as a development root certificate. On Windows PowerShell:

```powershell
certutil -addstore -user Root .certs\localhost.crt
```

Trust only certificate generated locally for this project. Remove it after development through `certmgr.msc` if no longer needed.

Start HTTPS development server:

```bash
npm run dev
```

Verify `https://localhost:3000/taskpane.html` loads without certificate warning. Use `manifest.local.xml` for local sideloading.

## Sideload

### Excel for web

1. Open working workbook copy.
2. Select **Home > Add-ins > More Add-ins**.
3. Open **My Add-ins**.
4. Select **Upload My Add-in**.
5. Upload `manifest.xml` for GitHub Pages, or `manifest.local.xml` while local server runs.
6. Open **Test Evidence Inserter** task pane.

### Excel on Windows

Use organization-approved network catalog:

1. Put `manifest.xml` in shared catalog folder. Use `manifest.local.xml` only for local development.
2. In Excel, open **File > Options > Trust Center > Trust Center Settings > Trusted Add-in Catalogs**.
3. Add catalog URL/path and mark it in menu.
4. Restart Excel.
5. Open working workbook copy.
6. Select **Home > Add-ins > More Add-ins > Shared Folder**.
7. Choose **Test Evidence Inserter**.

Some Microsoft 365 builds also expose **Upload My Add-in** directly.

## Use

1. Open test scenario sheet, or select scenario row on `Summary`.
2. Task pane refreshes automatically when worksheet tab changes. Use refresh icon if needed.
3. Choose Test Type: `General` provides only Generic Screenshot; `SOA` provides all SOA evidence types.
4. Choose Evidence Type.
5. For Generic Screenshot, choose **Left (A:D)** or **Right (F:I)** and optionally enter dynamic Header. Right images start at column F and match the `A:D` width.
6. To copy a picture from another workbook, select it in workbook A and press `Ctrl+C`. In workbook B, paste the Excel object onto the destination scenario sheet.
7. Select **Use pasted Excel picture**, then click the picture you just pasted. The task pane previews the picture.
8. Select **Insert**. After successful insertion, the temporary pasted object is replaced by the managed evidence image. The original in workbook A stays in place. Screenshot insertion never changes worksheet Status.
9. In separate **Worksheet Status** panel, choose `PASS/FAIL`, `PASS`, or `FAIL`. Changing dropdown immediately updates workbook; no button required.

Changing selected row while already on `Summary` does not activate another worksheet. Use refresh icon after changing Summary row.

Images are appended below existing evidence and user-owned content in `A:I`. Existing `Picture N` objects are not changed unless the tester explicitly selects a temporary pasted object through **Use pasted Excel picture**; that temporary object is removed only after its managed evidence copy is inserted. Values, formulas, and unmanaged shapes create protected vertical bands with one blank worksheet row after them. Earlier managed evidence stays fixed while the new screenshot continues below the protected content.

Use **Convert Existing** to bring worksheet evidence that predates the add-in under management. Select one cell or one row in `A:D` or `F:I`, then choose either **Header and pictures** to review and convert detected Excel picture objects or **Header for next insertion** to adopt only the header once. The review defaults to the evidence type inferred from the header and allows a manual override. Only checked pictures are converted.

Managed placement treats values and formulas as user-owned content, including formulas that display an empty string. Content anywhere in `A:I` blocks that vertical band, even when the opposite screenshot slot is empty. Paired screenshots remain aligned below the obstruction; only the affected pair row and required later rows in that group may move. Edited managed headings remain untouched. Empty table cells and formatting by themselves do not block placement.

Test Type, Status, and each profile's last Evidence Type, Generic Side, and Header are stored in Excel workbook settings. A workbook without saved choices defaults to `General` and `PASS/FAIL` Status. Changing Test Type restores that profile's choices while keeping pending screenshot and Status. Scenario refresh preserves Evidence Type, Side, and Header; recognized worksheet Status takes precedence. General and SOA evidence are isolated: one profile does not reflow or remove another profile's images.

### Test Scenario discovery

Worksheet names do not need a `TS` prefix. Task pane searches `A1:P12` for semantic labels such as `Test Case No`, `Condition`, and `Expected Result`, supporting horizontal or vertical layouts, multiline text, and data continuing through following rows. Numeric displayed IDs such as `11.0` remain text.

On `Summary`, task pane finds Scenario ID column from its header instead of assuming column A, then resolves worksheet named by selected row. Classification/title rows between headers and actual test data are ignored. Actual Test Case No, Condition, Expected Result, and Status stay bound to one row. Blank, utility, incomplete, or ambiguous sheets safely show `Not loaded` and disable workbook actions.

### Worksheet Status colors

Status update writes discovered Status cell and worksheet tab color together:

- `PASS/FAIL`: `#FFC000`
- `PASS`: `#00B050`
- `FAIL`: `#FF0000`

Blank or unrecognized worksheet Status displays persisted Status, defaulting to `PASS/FAIL`. Legacy stored `unchanged` values are coerced to `PASS/FAIL`.

Inserted image names follow:

```text
EV__SCENARIO_ID__MODULE_ID__SEQUENCE
EV__TS4_1_1__PDF_PAPER__001
```

**Remove Last Inserted** deletes only highest-sequence `EV__...` image for selected Test Type. Remaining managed groups compact upward while preserving pair alignment and staying below non-add-in shapes and user cell content appended below the managed block. Backward-compatible cleanup also removes matching legacy `EVC__...` caption shape if present.

New evidence headings use worksheet cells, not shapes:

- Left cell: column `A`
- Right cell: column `F`
- Gray fill, bold black text
- Images fit left `A:D` or right `F:I` slot; right images start at F and match the `A:D` width
- Every image spans its assigned slot width; height follows original aspect ratio without cropping
- MCD and Statement Duplicate intentionally have no heading
- Header input appears only for Generic Screenshot

Configured evidence partners share one two-column group:

- PDF pairs with AFP; PDF - Paperless pairs with AFP - Paperless; SOA Batch pairs with SOA Batch - Paperless
- Partner headings and screenshots align by ordinal row even when another evidence family was inserted between them
- The family remains at its first insertion position, so partner alignment takes priority over strict insertion order for paired evidence
- Extra screenshots remain under the shared headers on their configured side

Generic Screenshot groups use exact Header matching:

- Same exact text shares one gray heading cell
- Matching text rejoins its existing managed group even after other Headers were inserted (`A → B → A` displays one `A` block, then `B`)
- Capitalization and whitespace differences create separate groups
- Empty values add to most recently used Generic group without creating a header
- Empty value with no previous Generic group creates an unheaded group
- Matching images stack vertically without screenshot-count limits
- Group blocks follow first insertion; images inside each block follow insertion sequence. Visual order can therefore differ from global insertion chronology
- Legacy metadata upgrades on the next successful insert/remove without merging historical duplicate Header occurrences; a new match targets the most recently used matching legacy group

Generic Screenshot Side is explicit per insertion:

- Left fills `A:D`; Right starts at F and matches the `A:D` width
- Left followed by Right in the same exact Header group shares one row, even when another Header was inserted between them
- Consecutive same-side images start separate rows
- Taller paired image determines where the next row begins
- Existing Generic screenshots without Side metadata migrate to Left
- Side defaults to Left, then persists separately for General and SOA
- Header remains after insertion; use compact **Clear** button to remove and persist an empty Header

## Test

```bash
npm run typecheck
npm test
npm run build
```

### Normal SOA acceptance

1. Open working copy and `TS4.1.1`.
2. Insert PDF screenshot.
3. Insert AFP screenshot.
4. Confirm left/right images share top, retain aspect ratio, and do not overlap.

### Paperless acceptance

1. Open working copy and `TS4.25.1`.
2. Insert SOA Batch, then SOA Batch - Paperless.
3. Insert PDF, then AFP.
4. Insert PDF - Paperless, then AFP - Paperless.
5. Confirm each pair shares top and each later group starts below prior taller image.
6. Confirm all original screenshots remain unchanged.

## Security controls

- PNG/JPEG MIME allowlist
- Binary signature validation prevents MIME spoofing
- 10 MB maximum before Base64 conversion
- Browser decode validates image content and dimensions
- Workbook text rendered as React text, never raw `innerHTML`
- Generic Header text is escaped before writing to a cell; formula injection is blocked
- Non-generic evidence accepts no caption/header input
- Status constrained to static allowlist
- Screenshot Base64 held only in memory
- Pasted workbook picture previews are held only in memory
- No screenshot upload, backend, local storage, or credentials
- Existing shapes excluded from deletion unless name exactly matches managed pattern
- Heading writes stop when target cell contains user-owned content
- HTTPS required in development and production

## GitHub Pages deployment

Production manifest targets:

```text
https://thrddqno.github.io/screenshot_inserter/taskpane.html
```

Build and commit static bundle:

```bash
npm run build
git add docs manifest.xml
git commit -m "build(pages): update static bundle"
git push
```

Repository configuration:

1. Open **Settings > Pages**.
2. Select **Deploy from a branch**.
3. Select branch `main` and folder `/docs`.
4. Save and enable **Enforce HTTPS** when available.
5. Verify task-pane URL before sideloading `manifest.xml`.

GitHub Pages serves only add-in code. Screenshots stay inside Excel. Site and source can be public; never commit workbooks, customer data, internal URLs, credentials, tokens, or internal configuration.

Company rollout should use approved internal HTTPS hosting or private GitHub Enterprise Pages.

## Current limitations

- Only first milestone layouts are validated in configuration.
- Real Excel host acceptance must run on Windows Excel and Excel for web.
- Excel copies worksheet pictures as Office drawing objects that task-pane clipboard APIs cannot read directly. Cross-workbook transfer therefore requires a temporary native Excel paste onto the destination worksheet before selecting **Use pasted Excel picture**.
- Manual position changes to add-in-managed images are normalized during the next insertion or removal; unrelated shapes remain untouched.
- Evidence count is not updated because no evidence-count cell is configured.
