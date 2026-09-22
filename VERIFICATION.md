# Verification

The original release was checked against the two supplied ZTE ITBBU V5.85.20.20 exports for LG0068 and LG0011. The original measurements and earlier compatibility checks are retained below. The final section covers the current **v1.2.1** theme and Excel-metadata fixes using the latest, resaved workbooks from the error report. Earlier measurements describe earlier versions of the inputs, not the newest pair.

## Merged workbook

| Check | Result |
| --- | --- |
| File-1 configuration records | 17,061 |
| File-2 configuration records | 25,335 |
| Final configuration records | 42,396 |
| Configuration sheets verified | 1,899 |
| Total workbook sheets | 1,902 |
| Header cells preserved across rows 1–5 | 188,110 |
| Dropdown validation rules preserved | 7,720 |

Every output data row was checked against the source row, including all values, cell styles and file order. The original header-row XML is identical. TemplateInfo, Index and HideEnum are byte-identical to File-1. Other original package parts, including the stylesheet, are unchanged. Worksheet metadata is unchanged apart from the used-range dimension and the appended data payload. The original LDN strings are retained.

## Parameter audit

| Status | Parameter comparisons |
| --- | ---: |
| Same | 163,323 |
| Different | 13,361 |
| Missing object | 401,827 |
| Total | 578,511 |

These counts use the documented cross-site normalization with identity parameters excluded. Missing-object counts count parameter comparisons, not unique missing objects.

An independent verifier checked all 578,511 audit rows, including 755,195 present source values and their original LDNs, against the source workbooks. It checked site/row traceability, status classification and complete coverage with no duplicated, missing or unexpected audit records. Each file’s original LDN is shown separately from the comparison key.

The full application audit export completed in approximately 20 seconds in the test runtime, after inspection and comparison. Device/browser speed and available memory will change this duration.

## Application checks

The original nine regression tests covered three-file merging, reordered data rows, exact template headers, blank versus zero, text versus numeric cells, literal marker text, duplicate keys, missing keys, root normalization, original LDN preservation, shared strings, missing worksheets, incompatible metadata, formulas, audit exports, and selected-sheet scope.

The browser-worker scripts were also executed in an isolated browser-like JavaScript context using both original workbooks. Script loading, inspection, comparison and the original-LDN preview passed. JavaScript syntax and HTML asset/element references were checked.

No deployment to a Vercel account, live ZTE import, or interactive browser UI test was performed. The separate delivered audit uses the same verified dataset with additional report formatting.

## Compatibility update — v1.1.0 (21 September 2026)

All 16 regression tests passed at this stage. They retained the original data/audit checks and additionally verified:

- Description differences allow merging while File-1 header XML stays identical.
- Three differently ordered column layouts preserve every value, shared string and numeric type under the correct destination parameter; full primary keys match regardless of column order.
- Appended cells use the destination column style when the source style table differs.
- Later-file lookup changes, or missing lookup sheets, preserve File-1 support sheets, named ranges and validation definitions. Lookup checks resolve shared strings and ignore XML/style-only differences.
- Additional sheets carrying template rules still block if their lookup or named-range dependencies differ.
- Parameter-set, type/range and key-rule conflicts still block and identify source/destination cells. Formula protection is retained.
- The audit's Template details sheet contains the differing header values and the action taken.

An independent openpyxl integration check generated three valid XLSX inputs containing 360 records, five template rows, reordered columns, different header guidance, different hidden lookup contents, dropdowns, custom styles, blank/zero/Boolean/text values and literal text beginning with `=`. The merged result was reopened with openpyxl and all 360 records matched the sources under File-1 column names. File-1 header XML, styles, TemplateInfo, Index, HideEnum and workbook named ranges were preserved; validation ranges extended correctly. The exported audit reopened successfully with 240 comparisons and original LDNs.

The updated worker handled inspection and audit messages in an isolated JavaScript context. JavaScript syntax checks passed. The two small 5G workbooks in the 21 September error screenshot were not attached, so their exact row-3/row-5 compatibility has not been verified. The screenshot alone cannot establish which header values changed. This release does not assert that all different template versions are compatible, and no live ZTE import or deployed-browser test was performed.

## Widest-sheet merge — v1.2.0 (22 September 2026)

All **24 current regression tests pass**, and all three application scripts pass syntax checks. The new tests verify:

- A 23-column File-1, 35-column File-2 and 30-column File-3 select File-2's template. Every value is checked under the correct parameter, including reordered columns beyond Z, text identifiers, numeric zero and blank missing parameters.
- Reference selection is independent for each MO. Different MOs select File-1, File-2 and File-3; an empty but wider template remains eligible. Equal widths keep the earliest file.
- A later reference replaces the existing MO sheet without duplicating its name. Header rows come from the reference while data rows retain file order.
- Missing parameter columns, present blank values, zeros, original LDNs and duplicate keys remain distinct in the audit. The exported coverage report identifies the reference and all source column counts.
- Later references retain their fonts, fills, borders, number formats, column styles, conditional formatting, dropdowns and scoped print ranges. Imported style indices are remapped, and different lookup contents receive separate hidden sheets and named ranges. File-1 support sheets remain unchanged.
- Self-closing blank cells do not absorb the next parameter's value. Unmatched source parameters, conflicting shared definitions, formulas and unsupported dynamic/external template references are reported before an invalid merge can be downloaded.

The app was run against the newly uploaded complete V5.85.20.20 exports. Inspection returned no blocking issues. The merged result contains **42,396 configuration records across 1,899 MO sheets**. A separate XML reader independently compared every source/output value and type in order, all **188,110 header cells**, and **7,720 validation rules**. TemplateInfo, Index and HideEnum and the other original package parts remained byte-identical to File-1. The audit still reports **578,511 comparisons**: 163,323 Same, 13,361 Different and 401,827 Missing object.

A separate mixed-width integration run used three valid XLSX files with 72 records across three MOs, selecting a different reference file for each MO. Its merge and audit were reopened read-only with openpyxl. All header values and formatting, data values/types, missing-column blanks, source LDNs, column widths, freeze panes, dropdown dependencies and **672 audit comparisons** were verified. This included a later reference with a different style table and conflicting hidden lookup contents.

The v1.2.0 worker successfully handled inspection, audit and merged-download messages in an isolated JavaScript context. No interactive browser, Vercel deployment or live ZTE network-manager import was performed. The actual small 5G files shown in the screenshots were not available; their unequal-column scenario is covered by the regression tests, while any additional definition conflicts would still be reported explicitly.

## Theme and Excel-metadata compatibility — v1.2.1 (22 September 2026)

The exact latest uploaded files, locally suffixed `(2).xlsx`, were tested in the screenshot order and the reverse order:

- Screenshot File-1: `RANCM-Custom-template-huabdaka-20260916145056-ITBBU-ITRAN-PNF_V5.85.20.20(2).xlsx` — 8,255,262 bytes; 25,335 configuration records; a newer Excel theme.
- Screenshot File-2: `RANCM-Custom-template-huabdaka-20260916144841-ITBBU-ITRAN-PNF_V5.85.20.20(2).xlsx` — 7,408,400 bytes; 17,061 configuration records; a legacy export without a theme part.

The previous version reproduced both reported blocking issues exactly. File-2 supplies the wider SubRack and Slot sheets, but its absent theme differed from File-1's newer theme. The update resolves imported theme font and colour references rather than requiring identical themes. Independent output parsing also exposed an unrelated namespace-transfer defect: newer Excel rows carried `x14ac:dyDescent` into an older worksheet without its namespace declaration. The update preserves those row-level namespace bindings and their markup-compatibility declarations, including when a prefix means something different in the destination.

| Check | Screenshot order | Reverse order |
| --- | --- | --- |
| Blocking compatibility issues | 0 | 0 |
| Configuration records | 42,396 | 42,396 |
| MO sheets / total sheets | 1,899 / 1,902 | 1,899 / 1,902 |
| SubRack reference / columns / records | File-2 / 15 / 2 | File-1 / 15 / 2 |
| Slot reference / columns / records | File-2 / 9 / 20 | File-1 / 9 / 20 |
| Header cells checked | 188,110 | 188,110 |

A separate XML reader verified every output data value, stored type, parameter mapping and file order against these source files. Base-reference header-row XML was identical. For the two later-reference templates, it checked all header values, resolved fonts/fills/borders, column widths and column styles. TemplateInfo, Index, HideEnum and the other unchanged base package parts were byte-identical. Every standard dropdown rule retained its reference formula. The supplied resaved base contains no standard validation rules; its merged output gains 4 rules from the two later references. The reverse-order output retains all 7,720 rules from the legacy reference templates. The app retains the selected reference's actual rules; it does not invent missing rules. Both ZIPs passed CRC checks, their worksheets parsed as XML, and both workbooks opened with openpyxl read-only. The legacy-base output carries the source's absence of a named default style, which openpyxl reports as a non-fatal warning.

The screenshot-order audit completed with **578,511 comparisons**: 163,300 Same; 13,361 Different; 401,827 Missing object; and **23 Missing parameter**. The new missing-parameter count reflects the narrower SubRack and Slot columns, not lost values in the merge.

All **29 regression tests** pass, along with syntax checks for all three application scripts. The five new tests cover modern-theme versus theme-less references in both orders; identical stylesheet XML used with different themes; major/minor fonts, RGB/tint, borders, conditional formatting and tab colours; custom indexed palettes; shared/inline rich text; implicit style-zero cells; and scoped newer-Excel row namespaces. Prior column-mapping, audit, lookup and real-conflict protections remain tested.

The current browser-worker scripts handled inspection, audit and merge for the entire screenshot-order pair in an isolated JavaScript context. This is not an interactive browser test, a Vercel deployment, or a live ZTE network-manager import. The ZIP update must still be committed to the user's repository and redeployed before the live application uses v1.2.1.
