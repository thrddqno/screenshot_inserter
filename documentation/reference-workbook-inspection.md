# Reference workbook inspection

Inspected read-only:

```text
/home/adioquino/Downloads/CHG0038673_Test Results_SOA_RN4.0.xlsx
```

SHA-256:

```text
6a1c92ae48bb019e248aad1b492fab26c71ebc66c8aa5f5223f1ec1728868aaa
```

## Structure

- 80 visible sheets: `Summary` plus 79 scenario sheets.
- `Summary` has 82 scenario rows (`A2:H83`).
- `TS4.30.1`, `TS4.30.2`, and `TS4.30.3` appear in `Summary` but have no worksheets.
- Scenario ID is in `A2`; Condition in `B2`; Expected Result in `C2`; Status in `D2`.
- No explicit Evidence Profile or workbook defined name exists.
- Workbook embeds unrelated OMEX task pane `wa104380862`; preserve it.

## Screenshot storage

- Screenshots are PNG drawing shapes with generic `Picture N` names.
- They use two-cell anchors and one-cell edit behavior.
- Shape extents can exceed worksheet used range. `TS4.25.1` reports `A1:H375`, while shapes extend near row 678.
- Safe placement must inspect shape rectangles, not only used range.

## `TS4.1.1` normal SOA

| Heading cell | Text | Existing images |
|---|---|---:|
| A4 | INVOICING | 1 |
| A40 | SOA Batch | 1 |
| A78 | PDF | 2 left |
| F78 | AFP | 2 right |

Total: 6 images. PDF/AFP pairs begin at matching vertical positions.

## `TS4.25.1` Paperplus

| Heading cell | Text | Existing images |
|---|---|---:|
| A4 | Invoicing | 1 |
| A35 | SOA Batch | 1 left |
| E35 | SOA Batch - Paperplus | 1 right |
| A69 | PDF - Paper | 4 left |
| F69 | AFP | 4 right |
| A375 | PDF - Paperplus | 4 left |
| F375 | AFP - Paperplus | 4 right |

Total: 19 images. Pair groups share top positions; later groups begin below taller images.

## Ambiguities

- `TS4.25.1` Condition/Expected Result do not say Paperplus. Existing headings or explicit profile are required.
- `TS4.23.1` has malformed duplicate `PDF - Paperplus` headings and lacks clear `AFP - Paperplus`; confirmation required.
- 46 sheets have screenshots but no evidence headings.
- `Retailer Statement Duplicate` mapping needs future product decision.
- Existing screenshot counts vary: normal reference has 2 PDF/AFP pairs; Paperplus reference has 4 pairs per PDF section.
- Heading capitalization and row positions vary.
- Existing generic shape names provide no reliable evidence type metadata.
