# Verification

Checked against the two supplied ZTE ITBBU V5.85.20.20 exports for LG0068 and LG0011.

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

Nine regression tests cover three-file merging, reordered data rows, exact template headers, blank versus zero, text versus numeric cells, literal marker text, duplicate keys, missing keys, root normalization, original LDN preservation, shared strings, missing worksheets, incompatible metadata, formulas, audit exports, and selected-sheet scope.

The browser-worker scripts were also executed in an isolated browser-like JavaScript context using both original workbooks. Script loading, inspection, comparison and the original-LDN preview passed. JavaScript syntax and HTML asset/element references were checked.

No deployment to a Vercel account, live ZTE import, or interactive browser UI test was performed. The separate delivered audit uses the same verified dataset with additional report formatting.
