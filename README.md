# ITBBU Config Studio

A complete browser app for merging ZTE ITBBU configuration workbooks and comparing parameter values across two or more files. No API keys, database, Python backend or server uploads are required.

Version **1.2.0** selects the sheet with the most parameter columns as the reference for each MO. A 23-column sheet can merge into a 35-column reference from any uploaded file. Values match by exact parameter name; absent parameters remain blank. The reference supplies the five template rows, column order, formatting and dropdowns. Data rows still follow File-1, File-2, File-3 order.

## Update an existing GitHub deployment

Extract `ITBBU_Column_Merge_Update.zip` and replace the matching files in your repository. Keep the `public` and `tests` folder structure; upload the extracted files, not the ZIP itself. Commit the changes and redeploy the connected Vercel project. Reload the app and select the workbooks again. The desktop sidebar shows `v1.2.0`; script URLs are versioned so the new inspection and merge logic loads together.

## Deploy to Vercel

1. Extract this ZIP. The project root contains `vercel.json`, `package.json` and the `public` folder.
2. Put the extracted project contents in a GitHub repository.
3. In Vercel, choose **Add New → Project**, then import that repository.
4. Use **Other** for Framework Preset. Leave the build and install commands empty. Set Output Directory to **public**. The supplied `vercel.json` already defines these settings.
5. Deploy. Open the resulting HTTPS URL and select your workbooks.

If your repository contains this project inside another folder, set that folder as the Vercel Root Directory.

Alternatively, from the extracted project root, run `npx vercel` for a preview deployment or `npx vercel --prod` for production. Sign in to your own Vercel account when prompted. These commands require Node.js and internet access on your computer.

Vercel references: [Project configuration](https://vercel.com/docs/project-configuration), [Deployment methods](https://vercel.com/docs/deployments).

## Use the app

1. Select two or more `.xlsx` workbooks. The displayed order defines File-1, File-2, File-3 and so on. Remove and re-add files to change that order. File-1 supplies the base workbook; each MO independently selects its widest reference sheet.
2. Select **Inspect workbooks**. Sheet coverage shows each file's row and column counts and the selected reference for every MO. Review template notices and any blocking issues.
3. Choose a comparison mode and optionally select particular sheets to audit.
4. Select **Run parameter audit**. Search by sheet, parameter, object or value, or filter the status.
5. Download **Merged workbook** and **Audit workbook** separately. The audit export includes all selected sheets; preview search filters do not restrict the export. The explicit export checkbox restricts it to differences and issues.

The merged download always includes every configuration sheet, even when a subset is selected for auditing.

## Merge rules

- For each MO, select the sheet with the most parameter columns across all files, including empty sheets that carry a complete template. Equal counts use the earliest file. Keep that reference's rows 1–5 and original column order. Data begins on row 6. The reference may differ for every MO.
- Configuration sheets are identified by `MODIND` in A1. Parameter names in row 1 must be unique and nonempty. Match exact names, regardless of column position. A narrower source fills only the corresponding reference columns; its missing parameters remain blank, including missing mandatory/identity columns. No defaults or identifiers are invented.
- For example, with File-1 = 23 columns, File-2 = 35, and File-3 = 30, File-2 supplies the 35-column template. File-1 data comes first, with blanks in its 12 absent parameters, then File-2 and File-3 data. An unmatched parameter that exists in a source but not in the widest reference still blocks merging: no source column is silently discarded or added to the chosen template.
- For parameters present in both sheets, differences in descriptions/guidance in rows 2 and 4 are notices. Row 3 type/range/default definitions and row 5 primary-key/requirement rules must agree (apart from outer whitespace/line endings; row 5 also ignores case and repeated whitespace). Conflicts identify the parameter, source and reference cells, and both values. Different column counts alone do not block merging.
- Append all nonempty data rows in file order, preserving their values and Excel cell types. No deduplication or “last file wins” update is performed. Repeated source rows remain repeated.
- Preserve empty configuration sheets and the union of sheets found in the files. A wider later-file reference replaces the corresponding MO template in File-1's existing sheet position; it does not create a duplicate MO sheet. Sheets absent from File-1 are appended.
- Retain TemplateInfo and Index unchanged from File-1 and exclude them from merge/audit calculations. The Index is not regenerated.
- Preserve File-1 non-configuration support sheets, including hidden `HideEnum`, unchanged. These contain dropdown lookup data, not row-6 site records. If a later reference needs a different lookup, retain it under a separate hidden name such as `__ITBBU_F2_HideEnum` and redirect that reference's dropdowns/named ranges. Identical lookup contents are reused after resolving shared strings. Values are not translated between enum codes or validated against every vendor-specific rule.
- Preserve the reference's fonts, fills, borders, column widths, template row heights, print ranges and validation rules. Import its styles with remapped indices when necessary. Existing validation ranges extend for appended records. Source data from a different style table uses the reference's column/data styles. File-1 keeps its original support sheets, styles and named ranges; imported styles and names are added as needed, and scoped names for a replaced MO come from that MO's reference.
- Do not add provenance or audit columns to the merged configuration template. Source traceability belongs in the separate audit workbook.
- No formulas are evaluated. Data formulas, merged data cells, tables, drawings, hyperlinks and comments on configuration sheets require manual handling and block the merge rather than being silently discarded. Convert formulas to values and remove unsupported objects from a copy when appropriate.

## Audit rules

### Across sites (default)

Match local `ldn`; otherwise use non-site primary keys or `moId`. The ManagedElement sheet is a site-level singleton. Never match arbitrary row positions.

Root normalization is enabled by default. For known ENB/GNB CUCP, CUUP and DU function roots, only the final suffix equal to that row’s ManagedElement is replaced by `{site}`. For example:

```
ENBCUCPFunction=621-20_10068,CULTE=1,Cell=11
ENBCUCPFunction=621-20_10011,CULTE=1,Cell=11
```

Both become `ENBCUCPFunction=621-20_{site},CULTE=1,Cell=11` for matching. Cell IDs, neighbor IDs, PLMN and other path components remain exact. Disable normalization to require the original LDN string to match.

This comparison assumes equal local identifiers refer to comparable objects. That can be inappropriate for hardware positions, automatically created neighbors, or sites with different cell numbering. Review those keys before drawing conclusions. The app does not infer topology or identify a “correct” parameter value.

If several sites or rows in one file share a local key, the app reports **Ambiguous key**, includes the candidate values and source row numbers, and does not select an arbitrary match.

### Same site snapshots

Use all **Primary Key** columns marked in row 5, including site identity. Choose this mode when comparing successive exports of the same network elements. Root normalization does not apply in this mode.

### Parameter values and statuses

By default, MODIND, ManagedElementType, SubNetwork, ManagedElement, NE_Name, ldn and moId are excluded as comparison parameters. They still support matching and traceability. Enable **Include identity parameters** to compare them too.

| Status | Meaning |
| --- | --- |
| Same | The object is unique in every file and the parameter’s exact value/type matches. |
| Different | The object is unique in every file and at least one parameter value/type differs. |
| Missing object | At least one file has no row with that object key. |
| Missing sheet | The sheet does not exist in at least one file. |
| Missing parameter | The parameter column is absent in a contributing sheet. |
| Ambiguous key | More than one row in a file has that comparison key. All candidates are shown. |
| No key | The row has no usable object identifier. It is reported without pairing it by position. |

Statuses use the order No key → Ambiguous key → Missing sheet → Missing object → Missing parameter → Different → Same when multiple conditions apply. The file-value columns retain the specific missing labels.

Empty values are `(blank)`; zero is a real value. Text matching is case-sensitive and whitespace-sensitive. Numeric, Boolean, error and formula cells carry type labels. Actual source text equal to a special missing/blank marker is prefixed `[text]` to distinguish it. Reports write cell text safely, including text beginning with `=`.

The audit workbook contains Summary, Files, Sheet coverage and one combined Parameter audit table. Columns include sheet, comparison key, parameter and status, followed by each file’s **original LDN, parameter value and source site/row**, then reason. Original LDNs remain unchanged in these columns; the normalized comparison key is separate. Missing parameter columns are reported as **Missing parameter**, distinct from a present empty cell. Sheet coverage includes each source's column count and the selected reference file. For very large reports, the table is split before Excel’s row limit. Template issues are included when present. A **Template details** sheet records reference selections, missing columns, header differences, column reordering and lookup notices, including the action taken.

## If compatibility still blocks merging

The original generic `Rows 1–5 differ` message did not distinguish missing columns or changed descriptions from changed parameter definitions. Version 1.2.0 handles subsets and reordered columns using the widest reference. For a remaining row-3 or row-5 conflict on a shared parameter, inspect the named cells in both source workbooks or the audit's Template details sheet and use compatible source exports. The app does not assume that different enum meanings, data types or primary-key rules are interchangeable.

`HideEnum` is retained from File-1. Do not delete it to work around a warning. A later reference's ordinary lookup data can be imported separately. Lookup formulas, linked objects, unsupported worksheet extensions, differing themes that affect imported styles, and incompatible custom colour palettes require manual handling and produce a specific issue.

## Limits and compatibility

- Standard, macro-free, unencrypted `.xlsx` workbooks only. Convert `.xls`, `.xlsb`, `.xlsm` and password-protected files first.
- 2–20 files; 150 MiB combined compressed input; 600 MiB combined expanded input; 96 MiB per expanded ZIP entry.
- At most 2 million parameter comparisons per audit. Select fewer sheets for larger exports.
- Excel’s 1,048,576 rows per worksheet, 16,384 columns and 32,767 characters per cell still apply. Overlong duplicate-value cells produce an explicit error rather than being truncated.
- Intended for a current desktop Chrome, Edge or Firefox browser. Available device memory determines practical capacity. A full telecom export can contain hundreds of thousands of audit rows and take time to download.
- Processing runs in a Web Worker. Cancel terminates it; inspect the retained file selection again to restart. Closing the page discards the session. No account, telemetry, remote file storage, CDN dependencies or network API calls are used by the app.
- The application ZIP contains code only; the supplied telecom workbooks and generated results are delivered separately.

## Run locally

From the project root, run:

```
python -m http.server 8000 --directory public
```

Open `http://localhost:8000`. Do not open `index.html` by double-clicking it: Web Workers require an HTTP/HTTPS origin. Python is only a convenient local static server; Vercel does not need Python.

## Verification and source

Run `npm test` (or `node --test tests/engine.test.cjs`) for the included regression tests, and `npm run check` for JavaScript syntax checks. No dependency installation is needed.

- `public/engine.js`: OOXML reader, template-preserving merge, audit matching and workbook export.
- `public/worker.js`: background processing and progress messages.
- `public/app.js`, `public/index.html`, `public/styles.css`: interface.
- `public/vendor/jszip.min.js`: bundled JSZip 3.10.1, with its license.
- `vercel.json`: static hosting and response headers.

The code has been checked against the two supplied 1,902-sheet exports. See `VERIFICATION.md` for the measured results and validation scope. It has not been deployed to your Vercel account or validated by importing the output into a live ZTE network manager.

Third-party dependency: [JSZip](https://stuk.github.io/jszip/) 3.10.1 (MIT or GPLv3; distributed here under MIT). The original license is retained in `public/vendor/JSZip-LICENSE.md`.
