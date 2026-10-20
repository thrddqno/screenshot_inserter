# Test Evidence Inserter

Proof-of-concept Excel task-pane add-in. Organizes screenshots after tester manually executes and validates tests.

Milestone scope:

- Normal SOA layout represented by `TS4.1.1`
- Paperplus layout represented by `TS4.25.1`
- No scenario classification from `.1`, `.2`, or `.3` suffixes
- No backend or external screenshot upload

## Safety first

Never open reference workbook for write testing. Create working copy:

```bash
cp --reflink=auto \
  '/home/adioquino/Downloads/CHG0038673_Test Results_SOA_RN4.0.xlsx' \
  '/tmp/opencode/CHG0038673_Test Results_SOA_RN4.0-POC.xlsx'
```

Reference SHA-256 at inspection time:

```text
6a1c92ae48bb019e248aad1b492fab26c71ebc66c8aa5f5223f1ec1728868aaa
```

Workbook files are ignored by `.gitignore`.

## Architecture

```text
src/config/evidenceModules.ts   Typed module catalog
src/config/layoutProfiles.ts    Image, workbook, and layout configuration
src/services/imageValidation.ts MIME, signature, size, and decode validation
src/services/layoutEngine.ts    Pure pairing and geometry engine
src/services/shapeNames.ts      Managed shape names and sequences
src/services/excelService.ts    Office.js workbook reads and writes
src/App.tsx                     One modular task-pane UI
manifest.xml                    Stable XML add-in manifest
manifest.local.xml              Local HTTPS development manifest
docs/                           GitHub Pages production bundle
```

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

1. Open scenario sheet, or select scenario row on `Summary`.
2. Task pane refreshes automatically when worksheet tab changes. Select **Refresh** if needed.
3. Choose Test Type. Currently only `SOA` is available.
4. Choose Evidence Type.
5. For Generic Screenshot, optionally enter dynamic Header.
6. Focus task pane and press `Ctrl+V` after Windows Snipping Tool capture.
7. Leave Status as `unchanged` unless explicit update is intended.
8. Select **Insert Evidence**.

Changing selected row while already on `Summary` does not activate another worksheet. Select **Refresh** after changing Summary row.

Images are appended below all existing non-add-in shapes. Existing `Picture N` objects are not moved, renamed, or deleted.

Test Type selection lasts for the current task-pane session. Changing it keeps the pending screenshot, resets Evidence Type to that profile's default, and clears Generic Header. Managed image metadata isolates each profile so future test types do not reflow or remove SOA evidence.

Inserted image names follow:

```text
EV__SCENARIO_ID__MODULE_ID__SEQUENCE
EV__TS4_1_1__PDF_PAPER__001
```

**Remove Last Inserted Evidence** deletes only highest-sequence `EV__...` image. Backward-compatible cleanup also removes matching legacy `EVC__...` caption shape if present.

New evidence headings use worksheet cells, not shapes:

- Left cell: column `A`
- Right cell: column `F`
- Gray fill, bold black text
- Images fit left `A:D` or right `F:I` slot
- Every image spans its complete slot width; height follows original aspect ratio without cropping
- MCD and Statement Duplicate intentionally have no heading
- Header input appears only for Generic Screenshot

Generic Screenshot groups use exact Header matching:

- Same exact text shares one gray heading cell
- Capitalization and whitespace differences create separate groups
- Empty values add to most recently used Generic group without creating a header
- Empty value with no previous Generic group creates an unheaded group
- Matching images stack vertically without screenshot-count limits

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

### Paperplus acceptance

1. Open working copy and `TS4.25.1`.
2. Insert SOA Batch, then SOA Batch - Paperplus.
3. Insert PDF, then AFP.
4. Insert PDF - Paperplus, then AFP - Paperplus.
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
- Removed evidence leaves managed heading cells in place; no unrelated content is touched.
- Evidence count is not updated because reference workbook defines no configured count cell.
