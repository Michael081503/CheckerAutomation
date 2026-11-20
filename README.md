# CHCODE Database Checker

A Streamlit app for validating rows in Excel or CSV files against one or more databases. Upload your databases and files, review the detected layout, and click **Check files**. Every file is checked independently against every database. The app automatically detects headers, finds the lookup column, compares supported fields, and provides one Excel report download. It supports SALAD master lists, PARKED and LEGACY assignment templates, unfamiliar exports with reliable account identifiers, CHCODE campaign files, HSBC mobile blasts, DataGrid email blasts, SALAD and PL endorsement email blasts, PL and SALAD Viber blasts, and files without headers. It is prepared for Streamlit Community Cloud.

The interface uses a dark theme with compact upload panels, a teal check button, and results that expand when you need details. File setup is automatic; no advanced settings or column selections are needed.

Results use the full browser width, with more room for long filenames in **File overview** beside **Issues across files**. Tables use readable text at normal zoom and stack on smaller screens. Turn on **Screenshot view** next to the report download to hide uploads and detailed file cards and reduce the outer margins further. Turn it off to return to the full view; uploads, results, filters and the Excel report are preserved. The overview shows every file/database comparison and up to 12 issue entries, with source row numbers, lookup values and an explicit count when more issues are available in the report.

After checking, setup and detection explanations disappear. The overview shows each difference on separate **Your file** and **Database** lines, with long values wrapping and empty values labeled **(blank)**. Multiple differences in a row are separated for easier review, and issues stack into labeled cards on small screens. Field names, mappings, skipped-row information and full differences remain available in the per-file details and Excel report. Replacing an upload clears the results and restores the detection summary.

## CIMB email files and masterfiles

Double-click **Start Checker.bat**. Upload the CIMB masterfile under **Add your databases**, add the campaign CSV under **Add files to check**, and click **Check files**. The automatic checker recognizes headers rather than filenames or dates. In a CIMB masterfile with an `ACTIVE` worksheet, it checks only that sheet; `POUT`, `WAIVER`, and the other worksheets do not supply replacement matches. A standalone worksheet with the same master columns is also supported.

| Campaign column | Master column | Check |
| --- | --- | --- |
| `primary_key` | `LOAN ACCOUNT NO.` | Full account number; leading zeros preserved |
| `email` | `EMAIL ADDRESS` | Email, ignoring letter case and outer spaces |
| `name` | `CUST_NAME` | Customer name, using the existing name comparison rules |
| `PRODUCT` | `PLACEMENT` | PERSONAL LOAN = PLOAN CO / PLOAN PCO; LAZPAY = LAZPAY CO / LAZPAY PCO |

With intact account numbers, `primary_key` identifies the database record. If any account keys use scientific notation, the whole file uses email to locate candidate records, while still comparing the full account number. Scientific-notation keys never pass, even when the contact fields agree or the master also contains a scientific-notation key. Rounded digits cannot be recovered from this CSV. Missing or blank lookup values remain **NOT FOUND** or **BLANK KEY**. All four fields must match the same database row; shared emails cannot combine values from different accounts, and an unmatched row never switches lookup columns.

`cc`, `bcc`, and `indexone_uid` have no confirmed mapping in this layout and remain **NOT CHECKED**. Additional campaign fields remain unchecked. The Excel report preserves original values, lists each difference, records the lookup and rules, and includes a **Field issues** worksheet with counts by field. Scientific-notation warnings remain visible after checking and in Screenshot view, and are saved in the report's Summary. No campaign eligibility or waiver calculation is checked.

The supplied `1120 CIMB.csv` has **20,175 rows**. All emails occur in the masterfile's **ACTIVE** sheet, with **104 customer-name differences** and **no product differences** among the candidate records. All **20,175 account keys** use scientific notation and contain only **three distinct values**, so **no rows are fully verified**. Many name differences also show encoding damage. Re-export full account numbers as text and names with a consistent encoding, then run the check again. Email matches alone cannot establish the correct replacement account number.

The report is `reports/CIMB_check_2026-11-20.xlsx`; SHA-256 verification in `reports/CIMB_source_verification_2026-11-20.json` confirms that both inputs were unchanged.

Run the same automatic check from PowerShell:

```powershell
$inputFolder = Join-Path $env:USERPROFILE 'Downloads'
.\.venv\Scripts\python.exe scripts\check_files.py --auto --database (Join-Path $inputFolder 'CIMB MASTERFILE 11182026.xlsx') --files (Join-Path $inputFolder '1120 CIMB.csv') --output 'reports/CIMB_check_2026-11-20.xlsx'
```

## Database rows with blank lookup values

Database rows with a blank lookup identifier are automatically excluded from comparison. `None`, empty strings and whitespace-only IDs do not block a mixed database containing valid records. Entirely empty rows are already ignored. Other blank fields do not remove a customer record: a valid account with a blank email is still checked. The source workbook is read without modification.

Exclusions use each file's actual database lookup column, such as `Old I.C.`, `CHCODE`, account number, `Debtor ID` or phone. A blank `Old I.C.` does not discard a record used by a valid phone or debtor-ID lookup. The interface shows the skipped count, the Excel Summary records it, and a `Skipped DB` worksheet lists the original source row numbers and lookup column for every exclusion. Full-copy checks exclude those rows from missing-record counts too.

Campaign rows with blank or unknown lookup values remain **BLANK KEY** or **NOT FOUND**. A database without usable lookup values still receives an error. For unfamiliar layouts, populated database IDs must remain unique; blank IDs never establish customer identity or field-mapping evidence.

The November 13, 2026 `DataGrid.xlsx` contains one excluded row with a blank `Old I.C.` at source row 474. All **472 campaign rows** in the three supplied CBS APD files match across the six mapped fields: **25** in Bucket 4, **48** in Bucket 7 and **399** in Bucket 8. The report is `reports/CBS_APD_blank_key_check_2026-11-13.xlsx`. SHA-256 checksums confirmed the workbook and all three CSV inputs were unchanged.

## HSBC SMS Credit Seeker and OffUs Payer discount blasts

Upload `HSBC DB as of Nov 12, 2026.xlsx` under **Add your databases**, add both SMS blast workbooks under **Add files to check**, then click **Check files**. The checker detects this layout by headers, so future dates, renamed files, reordered columns and equivalent CSV exports work automatically. Each file is checked independently against the uploaded DB. No manual field mapping is needed.

| SMS file column | DB column | Check |
| --- | --- | --- |
| `ChCode` | First `CHCODE` column | Customer lookup; leading zeros preserved |
| `Mobile` | `Mobile phone` | Equivalent Philippine mobile formats accepted |
| `ChName` | `Name` | Customer name |
| `LastName` | `Last Name` | Existing HSBC rule ignores leading Mr./Ms., Mr. and Ms. titles; surname and suffixes remain checked |
| `ProductType` | `PROD TYPE` | Credit Card = CARDS; Personal Loan = LOAN |

All five fields must match the same DB row. A phone belonging to another account is **DIFFERENT**, even if it exists elsewhere in the DB. Repeated CHCODEs cannot combine values from different rows. Unknown CHCODEs are **NOT FOUND**, blank CHCODEs are **BLANK KEY**, and blank campaign mobile numbers need review. Missing required DB fields produce an error. The older `formatted mobile` layout remains supported.

The supplied files also contain `discount` and populated columns G–I without headers. These are retained as `Unnamed column G`, `Unnamed column H` and `Unnamed column I`. The checker validates the supported customer fields and marks the other four fields **NOT CHECKED** in the interface and report. There is no confirmed discount field or calculation rule in the supplied DB, and unnamed helper values do not establish a mapping. Excel formulas are read using their saved values; external workbooks are not opened and formulas are not recalculated. Other extra campaign fields remain unchecked unless their mapping is supported.

The supplied November 12, 2026 files contain **1,250 data rows**. Credit Seeker has **226 of 226 matching rows**. OffUs Payer has **1,020 of 1,024 matching rows** and **four mobile differences**, at source rows **386, 462, 707 and 954** (header is row 1). There are no missing or blank CHCODEs. These results cover the five mapped fields; they do not validate discount offers or campaign eligibility. The report is `reports/HSBC_SMS_discount_check_2026-11-12.xlsx`. SHA-256 checksums confirmed all three source workbooks were unchanged.

Run the same automatic check locally:

```powershell
$inputFolder = Join-Path $env:USERPROFILE 'Downloads\MAKERCHECKER'
.\.venv\Scripts\python.exe scripts\check_files.py --auto --database (Join-Path $inputFolder 'HSBC DB as of Nov 12, 2026.xlsx') --files (Join-Path $inputFolder '12-11-26 HSBC SMS BLAST Credit Seeker - With Discounts - Copy.xlsx') (Join-Path $inputFolder '12-11-26 HSBC SMS BLAST OffUs Payer - With Discounts - Copy.xlsx') --output 'reports/HSBC_SMS_discount_check_2026-11-12.xlsx'
```

## Viber discount batches with a standard DataGrid export

Upload `DataGrid(1).xlsx` under **Add your databases**, add the eight `VIBER_discount_HIGH_BAL...` and `VIBER_discount_LOW_BAL...` CSVs under **Add files to check**, then click **Check files**. The checker recognizes the columns rather than dates, filenames, batch numbers, or filename row counts. Renamed files, future batches, reordered columns, extra campaign columns, and equivalent Excel exports work through the same automatic upload workflow.

This DataGrid contains `Last Called No.` and `Balance` rather than `Phone`, `OB`, and `Level`. Its fixed mappings are:

| File column | DataGrid column | Comparison |
| --- | --- | --- |
| `Phone` | `Last Called No.` | Phone; customer lookup |
| `FullName` | `Name` | Customer name |
| `OB` | `Balance` | Numeric; commas and equivalent decimals accepted |

All three values must match the same database row. Equivalent Philippine mobile formats are accepted and leading zeros are preserved in the original values. A wrong customer name or balance produces **DIFFERENT**; an unknown phone produces **NOT FOUND** and a blank campaign phone produces **BLANK KEY**. Shared phone numbers cannot combine fields from different customer rows. The checker always uses `Last Called No.` and `Balance` for this export; it does not switch to `Email` or `Amount Referred` to make a value pass.

`FirstName`, `LastName`, and `CHCode` are staff/template fields in this layout. `Monthly`, `Terms`, `Promo`, and `Level` have no confirmed database mapping or calculation rule and remain **NOT CHECKED**. The app and report disclose these seven unchecked fields. The other DataGrid export with `Phone`, `OB`, and `Level` retains its four established checks, described below.

The supplied November 10, 2026 files contain **16,833 campaign rows**. **16,830 match the three supported fields; 3 have phones not found in `Last Called No.`**. There are no differing mapped values among found phones and no blank campaign phone keys. The affected source rows, including the header as row 1, are:

| Campaign | Source rows needing review |
| --- | --- |
| LOW BAL, Level 1, Batch 1 (2,515 rows) | 482 |
| LOW BAL, Level 5, Batch 1 (4,000 rows) | 2680, 3623 |

The report is `reports/SBF_VIBER_discount_check_2026-11-10.xlsx`. All nine input files were verified unchanged using SHA-256 checksums. The report includes every original row, its result, field mappings, and unchecked-field explanations.

Run the same automation locally:

```powershell
$inputFolder = Join-Path $env:USERPROFILE 'Downloads\AUTOCHECKER'
$campaignFiles = @(Get-ChildItem -LiteralPath $inputFolder -Filter 'VIBER_discount*.csv' -File | ForEach-Object { $_.FullName })
.\.venv\Scripts\python.exe scripts\check_files.py --auto --database (Join-Path $inputFolder 'DataGrid(1).xlsx') --files $campaignFiles --output 'reports/SBF_VIBER_discount_check_2026-11-10.xlsx'
```

Supported uploads are `.csv`, `.xlsx`, and `.xlsm`. Other layouts need a supported profile or reliable identifier evidence for automatic inference; unsupported or ambiguous data is flagged for attention. Fields without a confirmed mapping remain unchecked.

## DataGrid and SALAD ML together

Add both `DataGrid(1).xlsx` and `SALAD ML.xlsx` under **Add your databases**. Add `09-11-26 SBF SALAD EMAIL BLAST NEGRES.csv` under **Add files to check**, then click **Check files**. Future exports and renamed files with the same headers work automatically. A single database can still be checked on its own.

The campaign is checked against each database separately, using one lookup column for the whole file. When both identifiers are identical, `primary_key` is used for lookup and both identifiers are compared. All six mapped values must match one customer record within each database, including when an identifier appears more than once:

| CSV column | DataGrid column | SALAD ML column |
| --- | --- | --- |
| `primary_key`, `indexone_uid` | `Old I.C.` | `UID Code/Old IC` |
| `accountnumber` | `Account No.` | `caseAccountNo` |
| `financingnumber` | `Financing/Card No.` | `caseBaseAccountNo` |
| `chname` | `Name` | `caseCustomerName` |
| `email` | `Email` | `caseCustomerEmail` |

Identifiers and account numbers preserve leading zeros. Existing customer-name comparison rules apply. `cc`, `bcc`, `agentcode`, `agentname`, and `last4` remain **NOT CHECKED** because there is no confirmed mapping for these fields in these databases. An unchecked field is not counted as a match.

The app labels each comparison with its database. The report contains a **Summary** with database names and lookup columns, separate row results, differences and field mappings for every database/file pair, and a **Combined** worksheet for each campaign. A combined row is **MATCHED** only if it passes every database; a missing identifier, differing field or failed comparison keeps it in review. A match in one database cannot override an issue in another. Invalid databases block the run until corrected.

SALAD ML's unnamed export column is retained without shifting the named columns. DataGrid's final count-and-sums footer is recognized when its count and all populated totals reconcile; the footer may contain `#N/A` in a contact cell. Other rows with blank lookup identifiers are automatically excluded and listed in the report.

The November 9, 2026 input has **16,056 campaign rows**. All 16,056 match DataGrid; **15,920 match SALAD ML and 136 have customer-name differences**. There are no missing or blank campaign identifiers. The combined result is 15,920 passing rows and 136 needing review. The report is `reports/SBF_SALAD_NEGRES_both_databases_2026-11-09.xlsx`. Source files are read without modification.

Run the same automation locally:

```powershell
.\.venv\Scripts\python.exe scripts\check_files.py --auto --database "..\AUTOCHECKER\DataGrid(1).xlsx" "..\AUTOCHECKER\SALAD ML.xlsx" --files "..\AUTOCHECKER\09-11-26 SBF SALAD EMAIL BLAST NEGRES.csv" --output "reports/SBF_SALAD_NEGRES_both_databases_2026-11-09.xlsx"
```

Multiple databases require `--auto`. The output cannot overwrite any database or campaign input. Browser uploads and automatic CLI runs use the same checker and report format.

## SBF HIGH / LOW BAL Viber blasts with DataGrid

Upload `DataGrid(3).xlsx` as the database and add all HIGH BAL CSV and LOW BAL Excel batches under **Add files to check**. Click **Check files**. Future dates, renamed files, new batches, reordered columns, and additional campaign columns use the same header detection; no filename list or manual mapping is required. Unfamiliar layouts still use the existing identifier-based inference and flag unsupported or ambiguous layouts for review.

This layout has `Level` without a debtor-ID column. Its DataGrid uses `Phone` instead of `Email`. The checker matches `Phone` to DataGrid `Phone`, accepting equivalent Philippine mobile formats, and compares all four fields on the same database row:

| File column | DataGrid column | Comparison |
| --- | --- | --- |
| `Phone` | `Phone` | Phone number; customer lookup |
| `FullName` | `Name` | Customer name |
| `OB` | `OB` | Numeric; commas and equivalent decimals accepted |
| `Level` | `Level` | Text |

The checker uses `Phone` consistently even when `Last Called No.` is also present. Wrong names, balances, or levels produce **DIFFERENT**; unknown phones produce **NOT FOUND** and blank phones produce **BLANK KEY**. For a shared phone, all fields must match one database row. It does not combine values from different customers. Missing required database fields produce an error. PARKED/LEGACY debtor-ID templates, PL Viber debtor-ID files, and SALAD Viber `Level Endo` files retain their own detection and mappings.

`FirstName`, `LastName`, and `CHCode` are staff/template fields in this layout. `Monthly`, `Terms`, and `Promo` have no confirmed database mapping or calculation rule. These six fields remain **NOT CHECKED**, and additional unknown columns also remain unchecked. All original values, field coverage, and differences appear in the Excel report. A mapped-row match covers only the four checked fields.

The supplied November 6, 2026 files contain 16,723 campaign rows: 964 HIGH BAL and 15,759 LOW BAL. All match across the four supported fields, with no differences, missing phones, or blank phone keys. The existing customer-name rule accepts an encoding replacement character only for a corresponding ñ. The local report is `reports/SBF_VIBER_BAL_check_2026-11-06.xlsx`; inputs are read without modification and reports remain excluded from GitHub.

Run the same automatic check locally:

```powershell
$campaignFiles = @(Get-ChildItem '..\AUTOCHECKER\06-11-26 SBF VIBER*' | ForEach-Object { $_.FullName })
.\.venv\Scripts\python.exe scripts\check_files.py --auto --database "..\AUTOCHECKER\DataGrid(3).xlsx" --files $campaignFiles --output "reports/SBF_VIBER_BAL_check_2026-11-06.xlsx"
```

## SBF SALAD email blasts with DataGrid

Upload `DataGrid (1)(1).xlsx` as the database and `05-11-26 SBF SALAD EMAIL BLAST.csv` as the file to check. Select **Check files**. Detection uses headers and identifiers, so renamed files and future exports with the same layout work without selecting columns. The website omits the static tutorial paragraphs below the upload panels; detection, results, field coverage, and report downloads remain available.

The SALAD email layout includes agent assignments, `OB`, `CRITERIA`, and three monthly/payment offers. Both `primary_key` and `indexone_uid` match `Old I.C.`; when they contain identical matching identifiers, the browser uses `primary_key` for lookup and checks both columns. Every mapped value must match the same database row:

| File column | DataGrid column | Comparison |
| --- | --- | --- |
| `primary_key`, `indexone_uid` | `Old I.C.` | Text; leading zeros preserved |
| `accountnumber` | `Account No.` | Text |
| `email` | `Email` | Text |
| `chname` | `Name` | Customer name |
| `agentcode` | `AgentCode` | Text |
| `agentname` | `AgentName` | Text |
| `OB` | `OB` | Numeric; commas and equivalent decimal formatting accepted |

These eight checks apply when the corresponding database columns exist. Wrong values remain differences; the checker never picks a different customer just to make a row pass. `cc`, `bcc`, `CRITERIA`, `monthly1`–`monthly3`, and `amount1`–`amount3` remain **NOT CHECKED**. Payment offers and campaign criteria have no confirmed database mapping or calculation rule. Their original values and the coverage are preserved in the report. Other SALAD endorsement and Viber layouts retain their own mappings.

The supplied November 5, 2026 pair contains 15,000 data rows. All 15,000 match across the eight mapped fields, with no differences, missing accounts, or blank keys. The local report is `reports/SBF_SALAD_EMAIL_check_2026-11-05.xlsx`; source data and reports are excluded from GitHub.

Run the same detection and generate a report locally:

```powershell
.\.venv\Scripts\python.exe scripts\check_files.py --auto --database "..\AUTOCHECKER\DataGrid (1)(1).xlsx" --files "..\AUTOCHECKER\05-11-26 SBF SALAD EMAIL BLAST.csv" --output "reports/SBF_SALAD_EMAIL_check_2026-11-05.xlsx"
```

## Start

From this folder, install the pinned dependencies and launch Streamlit:

```powershell
python -m pip install -r requirements.txt
python -m streamlit run streamlit_app.py
```

On Windows, **Start Checker.bat** launches the same app in your browser. It uses this folder's `.venv` when available, otherwise the `python` command. Keep the terminal open while using the local app. The Streamlit app requires Python 3.10+; it was tested on Python 3.13.

## PARKED and LEGACY templates

Check each template with the database that has the same name without `TEMPLATE`:

| Add your database | Add files to check | Supplied template rows |
| --- | --- | ---: |
| `PARKED.xlsx` | `PARKED TEMPLATE.csv` | 2,324 |
| `LEGACY.xlsx` | `LEGACY TEMPLATE.csv` | 1,470 |

Double-click **Start Checker.bat**, upload one pair, and select **Check files**. Download its report, then use **Start over** to check the other pair. Both supplied pairs were checked successfully: all 2,324 PARKED rows and 1,470 LEGACY rows match across the four checked fields, with no missing or blank debtor IDs. Headers are excluded. The local reports are `reports/PARKED_check_2026-11-03.xlsx` and `reports/LEGACY_check_2026-11-03.xlsx`. The checker validates four fields on the same debtor record:

| Template column | Database column | Check |
| --- | --- | --- |
| PARKED's unnamed column K, displayed as `Unnamed column K`; LEGACY's `Debtor` | `Debtor ID` | Customer lookup; identifiers preserve leading zeros |
| `FullName` | `Name` | Customer name |
| `LastName` | `Account No.` | Account number as text |
| `Phone` | `Email` | Phone number, accepting equivalent Philippine mobile formats |

The source headers are unusual: `LastName` contains an account number, and the database's `Email` column contains the intended phone number. These established template mappings remain fixed even when every uploaded value differs, so incorrect values are reported rather than causing a different mapping. Identifier checks never switch to another lookup column for an individual row. A found debtor with a wrong account, name, or phone is **DIFFERENT**; an unknown debtor is **NOT FOUND**, and a blank debtor is **BLANK KEY**. `FLOW_STAGE` is optional; the filename identifies each template.

`FirstName` and `CHCode` contain staff assignments and remain **NOT CHECKED**. `OB`, `Monthly`, `Terms`, `Promo`, and `Level` also remain unchecked. Blank or constant campaign columns are not evidence of a database match. All original values, including the unnamed column, remain in the report; the input files are read without modification.

Run the same automatic checks from this folder:

```powershell
.\.venv\Scripts\python.exe scripts\check_files.py --auto --database "..\AUTOCHECKER\PARKED.xlsx" --files "..\AUTOCHECKER\PARKED TEMPLATE.csv" --output "reports/PARKED_check_2026-11-03.xlsx"
.\.venv\Scripts\python.exe scripts\check_files.py --auto --database "..\AUTOCHECKER\LEGACY.xlsx" --files "..\AUTOCHECKER\LEGACY TEMPLATE.csv" --output "reports/LEGACY_check_2026-11-03.xlsx"
```

## New file and database layouts

The browser app and command-line `--auto` mode can inspect unfamiliar Excel and CSV layouts. The database needs a debtor, customer, or account identifier with unique populated values. The checker considers compatible identifier headers and matching values, then checks candidate fields against those same identified database records. Names alone, staff assignments, repeated campaign values, and blank/blank pairs cannot establish customer identity.

Review the detected lookup, proposed fields, and evidence before selecting **Check files**. Known profile mappings remain fixed so changed values still produce differences. For unfamiliar layouts, fields without enough header or same-record evidence stay **NOT CHECKED**. Unknown or ambiguous layouts receive an error requiring attention; they are not reported as passing. The selected lookup applies to the whole file, including unknown and blank identifiers. Future exports with the established PARKED/LEGACY layout are recognized from their columns rather than their filenames.

## Deploy to Streamlit Community Cloud

Upload the app source to a GitHub repository, then select that repository in Streamlit Community Cloud with **Main file path: `streamlit_app.py`** and **Python 3.13** in Advanced settings. See [DEPLOYMENT.md](DEPLOYMENT.md) for the exact file list and steps. The official [deployment guide](https://docs.streamlit.io/deploy/streamlit-community-cloud/deploy-your-app/deploy) describes the repository and entrypoint settings.

Upload your databases and campaign files through the app after deployment. The repository's `.gitignore` excludes Excel/CSV inputs, generated reports, uploads, and Streamlit secrets.

Uploaded data is processed on the app server. Each browser session keeps its own inspection metadata, results, and downloadable report. The app does not use a shared data cache or save uploads to the repository. Temporary files are deleted after inspection or checking, including failures. **Start over** resets the current app session. Uploads and results do not survive a server restart.

## Check your files

1. Add one or more database `.xlsx`, `.xlsm`, or `.csv` files in **Add your databases**.
2. Add multiple files in **Add files to check**. Excel and CSV files can be mixed, including files with and without headers.
3. Review the detection summary for each file's worksheet, header row, lookup column, and status. The database needs a `CHCODE`, `primary_key`, or `indexone_uid` header, supported DataGrid or endorsement headers, or a account identifier with unique populated values for automatic inference. If a workbook has multiple supported worksheets, save the intended worksheet as a separate file and upload it.
4. Click **Check files** to compare your uploaded rows against every database; smaller campaign subsets are supported.
5. Read the compact results, then expand a file to see its **Issues**, **Rows**, or **Fields checked**. Row previews default to entries needing review when issues exist.
6. Click **Download Excel report** for every original row, changed value, and field mapping.

The app selects the supported worksheet and lookup column automatically. Uploaded headers named `CHCODE`, `primary_key`, or `indexone_uid` can contain CHCODE, at any column position; keep the original header names. Recognized DataGrid databases use `Old I.C.` as the database key. It detects each uploaded file separately, so files with the same headers can use different lookup columns when their CHCODE values are stored differently. Recognized field mappings are applied automatically. **Fields checked** and the report show which fields were compared and which remain unchecked.

These header aliases are defined in the shared `checker.py` source and apply to the web app, desktop app, and command-line checker. They work with ordinary header layouts as well as recognized campaign templates. Explicit `CHCODE` takes priority in CHCODE layouts; recognized SALAD Viber blasts use phone and full name as described below. Existing DataGrid email layouts retain their usual defaults, and the web app uses database matches to distinguish candidate identifiers.

In ordinary CHCODE layouts, files without `FLOW_STAGE` that contain a named phone column compare each number with the database record for the same CHCODE. The detected `CHCODE`, `primary_key`, or `indexone_uid` column remains the lookup key, and other supported fields are also compared. Supported phone headers include `Mobile phone`, `formatted mobile`, `Phone Number`, `Mobile Number`, `Phone`, `Mobile`, and `Contact Number`; case, spaces, and underscores are ignored. Such files need a CHCODE lookup column and a supported database phone column. PARKED/LEGACY templates and inferred layouts use their disclosed debtor/account key and phone mapping instead. Files use their filenames when `FLOW_STAGE` is absent. A present `FLOW_STAGE` must contain one consistent, nonblank identifier. Files without headers use column B as CHCODE, retain the first row as data, and are identified by filename. The database does not need a `FLOW_STAGE` column.

## HSBC mobile blasts without FLOW_STAGE

Use `HSBC DB as of Oct 23, 2026.xlsx` as the database and add `23-10-26 BLAST Field.xlsx` and `23-10-26 BLAST NEW ENDO.xlsx` as files to check, then click **Check files**. Headers and comparison fields are detected automatically. The database does not need a FLOW_STAGE column.

When FLOW_STAGE is absent, the app checks each blast's `formatted mobile` against `Mobile phone` for the same CHCODE. A number belonging to another CHCODE is reported as a difference. Both files use their filenames as identifiers.

The mobile-blast layout compares these mapped fields for each CHCODE:

| File column | Database column | Comparison |
| --- | --- | --- |
| `formatted mobile` | `Mobile phone` | Phone |
| `ChCode` | `CHCODE` | Text |
| `ChName` | `Name` | Text |
| `LastName` | `Last Name` | Text |
| `ProductType` | `PROD TYPE` | Product |

Phone comparison accepts equivalent single Philippine mobile numbers in `+639`, `639`, `09`, and `9` formats, including spaces, hyphens, and parentheses. Different numbers, multiple numbers, and text with extensions are not merged. For account-based checks, product comparison accepts `Credit Card` / `CARDS` and `Personal Loan` / `LOAN`. In recognized HSBC mobile/SMS layouts, last-name comparison ignores a leading `Mr./Ms.`, `Mr.`, or `Ms.` title followed by whitespace. The surname and suffixes are still checked: `Lim Jr` matches `Mr./Ms. Lim Jr`, while `Lim` does not. Other fields, keys, other layouts, and full-copy checks retain their existing comparisons. Original titles remain in source values and reports.

For `28-10-26 HSBC SMS BLAST NEW ENDO.xlsx` and `28-10-26 HSBC SMS BLAST SMS II.xlsx`, upload `HSBC DB as of Oct 28, 2026(1).xlsx` as the database. The NEW ENDO file contains 68 data rows and the SMS II file contains 99. All 167 rows match across their mapped fields after ignoring the last-name titles, with no blank or missing CHCODEs. The extra `Mobile` column in SMS II is also checked against `Mobile phone`. `Blast Type` and `Endo` remain unchecked, and header rows are excluded. See `reports/HSBC_SMS_check_2026-10-28.xlsx`.

Successful checks without a stage show **MAPPED ROWS MATCH**. Wrong or blank phone numbers show **DIFFERENT**, including when the database phone is also blank. Unknown CHCODEs show **NOT FOUND**; blank CHCODEs show **BLANK KEY**. Repeated uploaded rows are checked independently, and duplicate database CHCODEs are compared using their mapped values. The app and Excel report show the expected number from the same account and disclose all field coverage. A present FLOW_STAGE column still needs one consistent, nonblank value. Explicit engine and command-line checks retain their existing FLOW_STAGE requirements, including full-copy comparison; the automatic app enables missing-stage checks with `Options.allow_missing_flow_stage` in uploaded-row mode.

The database and uploaded files can have titles, notes, or blank rows above their column headers. The app detects `CHCODE`, `primary_key`, `indexone_uid`, or the recognized DataGrid header layout, ignoring letter case and outer spaces, and skips the rows above it. The detection summary displays the worksheet, header row, and lookup column. Source row numbers in results and reports retain their original positions.

## SBF PL Viber blasts with DataGrid

1. Double-click **Start Checker.bat**.
2. Add `DataGrid(1).xlsx` in **Add your database**.
3. Add `27-10-26 SBF PL VIBER BLAST Level_1.csv`, `Level_3.csv`, and `Level_5.csv` in **Add files to check**.
4. Click **Check files**, review the differences, and download the Excel report.

These files are in `C:\Users\SPM\Desktop\AUTOCHECKER`. The layout is detected from its headers, so future exports with the same layout also work without changing filenames or adding FLOW_STAGE. A present FLOW_STAGE still requires a consistent, nonblank value.

PL Viber files contain `Level` and `Debtorid`. Their `CHCode` and `FirstName` values are staff codes and names. The checker uses the customer identifier instead:

| PL Viber column | DataGrid column | Comparison |
| --- | --- | --- |
| `Debtorid` | `Debtor ID` | Customer lookup; text identifiers preserve leading zeros |
| `FullName` | `Name` | Customer name on the same Debtor ID record |
| `Phone` | `Last Called No.` | Phone on the same Debtor ID record |

All three fields must match one database row. A phone belonging to another customer produces **DIFFERENT** even if that phone exists elsewhere in DataGrid. Equivalent Philippine mobile formats match, and the existing customer-name rule ignores question marks. A blank phone produces **DIFFERENT**, including when the database phone is blank. Unknown and blank debtor IDs produce **NOT FOUND** and **BLANK KEY**. Repeated uploaded records are checked independently. Duplicate database debtor IDs pass only when all mapped fields match one row.

`FirstName`, `LastName`, `CHCode`, `OB`, `Monthly`, `Terms`, `Promo`, and `Level` remain **NOT CHECKED**. The report preserves every uploaded value, lists the actual database lookup column per file, and shows file and database values for each difference. A passing result covers the three mapped fields. PL, SALAD, and ordinary CHCODE files retain their own lookup rules within a mixed batch.

The supplied database contains 2,342 records. The supplied PL Viber files contain 2,338 rows; every debtor ID was found:

| File | Rows | Matched | Rows needing review |
| --- | ---: | ---: | ---: |
| Level_1 | 148 | 120 | 28 |
| Level_3 | 111 | 102 | 9 |
| Level_5 | 2,079 | 1,932 | 147 |
| Total | 2,338 | 2,154 | 184 |

There are 23 name differences and 162 phone differences; one row has both. See `reports/SBF_PL_VIBER_check_2026-10-27.xlsx`. Uploaded-row checks allow campaign subsets, so unused database records do not count as missing from the campaign files.

The existing command-line checker detects this pair too:

```powershell
.\.venv\Scripts\python.exe scripts\check_files.py --database "..\AUTOCHECKER\DataGrid(1).xlsx" --files "..\AUTOCHECKER\27-10-26 SBF PL VIBER BLAST Level_1.csv" "..\AUTOCHECKER\27-10-26 SBF PL VIBER BLAST Level_3.csv" "..\AUTOCHECKER\27-10-26 SBF PL VIBER BLAST Level_5.csv" --output "reports\SBF_PL_VIBER_check.xlsx"
```

The five `28-10-26 SBF VIBER BLAST` CSVs use the same layout with `DataGrid(3).xlsx`. Their header rows are excluded from the 16,447 account rows checked. All Debtor IDs are present and unique across the batch. With the damaged-Ñ name rule described below, 16,372 rows match and 75 rows have phone differences; the 64 names containing a replacement character now match their account's DataGrid name. The rechecked report is `reports/SBF_VIBER_DebtorID_recheck_2026-10-28.xlsx`.

## SBF SALAD Viber blasts with DataGrid

Add `DataGrid salad.xlsx` as the database and add the `26-10-26 SBF SALAD VIBER BLAST Level_1.csv`, `Level_3.csv`, and `Level_5.csv` files to check. Click **Check files**. The app detects this format automatically without adding FLOW_STAGE or changing the source files.

The recognized layout contains `FullName`, `FirstName`, `LastName`, `Phone`, `CHCode`, `OB`, `Monthly`, `Terms`, `Promo`, and `Level Endo`. In these exports, `CHCode` contains repeated campaign labels rather than DataGrid account identifiers. This layout therefore uses these mappings:

| File column | Database column | Comparison |
| --- | --- | --- |
| `Phone` | `Last Called No.` | Phone lookup; equivalent Philippine mobile formats match |
| `FullName` | `Name` | Text, on the same database row as the phone |

Customer-name comparisons ignore `?` characters in both uploaded and database names. This applies to full names, `Name`, `ChName`, and first, middle, and last names, including mapped fields and full-copy checks. Existing case and space options still apply. Original names are preserved in source rows and Excel reports, and mapping rules disclose the comparison behavior.

Some exports already contain the Unicode replacement character `�` (U+FFFD) where a name had `Ñ` or `ñ`. On the same account record, `Pe�a` therefore matches `Peña` when every other character and mapped field agrees. This rule applies only to customer names: `�` does not match other letters or disappear from IDs, email addresses, or notes. Original values remain visible in the report. Ordinary `Pena` and `Peña` still differ.

Both mapped fields must match one DataGrid row to receive **MATCHED**. A matching phone with a different full name receives **DIFFERENT**, showing the expected name in the app and report. Unknown phones receive **NOT FOUND** and blank phones receive **BLANK KEY**. Shared database phones are resolved by the matching full name, and repeated uploaded rows are checked independently. The report summary includes each file's actual database match column.

`FirstName`, `LastName`, `CHCode`, `OB`, `Monthly`, `Terms`, `Promo`, and `Level Endo` remain **NOT CHECKED**. The campaign amounts are not treated as DataGrid's balance, monthly instalment, or paid term. A passing result covers phone and full name only. Other CHCODE layouts continue comparing phones for the same CHCODE, including in a batch containing both layouts. Recognized SALAD Viber pairs can omit FLOW_STAGE in uploaded-row checks, including the command line; a present stage and full-copy checks retain their usual validation.

The report for the supplied files is `reports/SBF_SALAD_VIBER_check_2026-10-26.xlsx`.

## CBS PL recovery and write-off email blasts

Upload `DataGrid(1).xlsx` as the database and add both `30-10-26 CBS PL EMAIL BLAST RECOVERY ACCOUNTS.csv` and `30-10-26 CBS PL EMAIL BLAST WRITE OFF ACCOUNTS.csv` as files to check. Click **Check files**. No column selections or FLOW_STAGE are needed; future exports with the same layout are detected from their headers.

These files use `primary_key` to find `Old I.C.`. The checker compares five fields on that same database record:

| File column | DataGrid column |
| --- | --- |
| `primary_key` | `Old I.C.` |
| `email` | `Email` |
| `accountnumber` | `Account No.` |
| `chname` | `Name` |
| `campaignbucket` | `Account Type` |

The supplied `indexone_uid` column is blank in every row. Automatic DataGrid email checks treat an entirely blank secondary identifier as unused and show it as **NOT CHECKED**, with the reason in **Fields checked** and the report's mapping sheets. `cc` and `bcc` also remain unchecked. The selected lookup key is always required. If a secondary identifier has any populated value, every row in that column is compared, including blanks and incorrect values. Explicit field mappings continue checking the fields selected by the user. This rule does not apply to ordinary CHCODE layouts or full-copy checks.

All 474 supplied campaign rows match the five checked fields. Their primary keys are unique across the two files and cover all 474 database accounts:

| File | Rows | Matched | Differences | Missing keys |
| --- | ---: | ---: | ---: | ---: |
| RECOVERY ACCOUNTS | 81 | 81 | 0 | 0 |
| WRITE OFF ACCOUNTS | 393 | 393 | 0 | 0 |

The local report is `reports/CBS_PL_check_2026-10-30.xlsx`. Inputs and generated reports stay outside Git; the repository includes synthetic regression examples.

To run the same check from this folder:

```powershell
.\.venv\Scripts\python.exe scripts\check_files.py --database "..\AUTOCHECKER\DataGrid(1).xlsx" --files "..\AUTOCHECKER\30-10-26 CBS PL EMAIL BLAST RECOVERY ACCOUNTS.csv" "..\AUTOCHECKER\30-10-26 CBS PL EMAIL BLAST WRITE OFF ACCOUNTS.csv" --output "reports\CBS_PL_check_2026-10-30.xlsx"
```

## DataGrid email blasts

For the four `29-10-26 CBS APD EMAIL BLAST` CSVs (BUCKET 4, BUCKET 7, BUCKET 8, and WRITEOFF), upload `DataGrid.xlsx` as the database and add all four CSVs. The checker uses `primary_key` to match `Old I.C.` and checks both identifier columns, email, customer name, account number, and campaign bucket on the same database record. `cc` and `bcc` remain unchecked. No FLOW_STAGE is needed.

This DataGrid export has an incorrect saved worksheet range (`A1:C1`). The Excel reader uses the actual worksheet cells, including all 37 columns. Its final count-and-sums row is excluded only when its count equals the preceding account count, all other populated cells are supported financial totals, and each total equals its column's sum. Real account rows with blank keys still fail validation. The same reader is used by the web app, desktop app, and command line, without editing the source workbook.

The supplied workbook contains 1,453 accounts. All 1,060 campaign keys are found: BUCKET 4 has 167 rows, BUCKET 7 has 65, BUCKET 8 has 486, and WRITEOFF has 342. Both identifiers, names, account numbers, and buckets match. DataGrid's Email cells are blank for all these accounts, so all 1,060 rows are reported as **DIFFERENT** for email. A complete database Email column is needed to validate the recipients. See `reports/CBS_APD_check_2026-10-29.xlsx` for source rows, database row references, field mappings, and differences.

To run the same check from this folder:

```powershell
.\.venv\Scripts\python.exe scripts\check_files.py --database "..\AUTOCHECKER\DataGrid.xlsx" --files "..\AUTOCHECKER\29-10-26 CBS APD EMAIL BLAST BUCKET 4 91-120.csv" "..\AUTOCHECKER\29-10-26 CBS APD EMAIL BLAST BUCKET 7 181-360.csv" "..\AUTOCHECKER\29-10-26 CBS APD EMAIL BLAST BUCKET 8 361+.csv" "..\AUTOCHECKER\29-10-26 CBS APD EMAIL BLAST WRITEOFF.csv" --file-key primary_key --database-key "Old I.C." --output "reports\CBS_APD_check_2026-10-29.xlsx"
```

To check `22-10-26 MSB PL EMAIL BLAST.csv`, add `DataGrid.xlsx` as the database, add the CSV as a file to check, and click **Check files**. No CHCODE or FLOW_STAGE columns need to be added.

To check `22-10-26 CBS APD EMAIL BLAST BUCKET 4 91-120.csv`, add `DataGrid (1).xlsx` as the database and the CBS APD CSV as a file to check. Click **Check files**. The checker detects the CBS file's `primary_key` column as its CHCODE lookup and compares it with the database's `Old I.C.`.

To check `22-10-26 CBS PL EMAIL BLAST.csv`, add its `DataGrid.xlsx` database and the CBS PL CSV, then click **Check files**. The checker tests the candidate identifiers against the database and detects which column contains CHCODE.

The checker recognizes a database with `Name`, `Old I.C.`, `Account No.`, and `Email`, paired with a blast containing `email`, `chname`, `accountnumber`, and `indexone_uid` or `primary_key`. When both identifiers are present, the app compares all their values against the database and uses the column with more matching keys. If both columns contain identical values and have matching keys, it uses `primary_key`, including the SBF layout. Other ties or files with no matching keys keep the layout's usual key: `primary_key` when `campaignbucket` is present, otherwise `indexone_uid` when available. A `CHCODE` header takes priority. The detection summary shows the selected column. Automatic comparisons are:

| Email-blast column | DataGrid column | Comparison |
| --- | --- | --- |
| `indexone_uid` | `Old I.C.` | Text |
| `primary_key` | `Old I.C.` | Text |
| `email` | `Email` | Text |
| `chname` | `Name` | Text |
| `accountnumber` | `Account No.` | Text |
| `financingnumber` | `Financing/Card No.` | Text, when both columns exist |
| `campaignbucket` | `Account Type` | Text, when both columns exist |

Both identifier columns are checked when populated, including an `indexone_uid` that differs from the selected `primary_key`. An entirely blank secondary identifier is disclosed as **NOT CHECKED** in automatic DataGrid email checks; partially populated columns and explicit mappings are still compared. Leading zeros in text identifiers and account numbers are preserved. Fields without a supported database counterpart, such as `cc` and `bcc`, remain **NOT CHECKED**.

For `23-10-26 SBF PL EMAIL BLAST.csv`, upload the corresponding `DataGrid.xlsx` database. The app detects `primary_key` as CHCODE and compares six fields when the database contains `Financing/Card No.`: both identifiers, email, name, account number, and financing number.

For this recognized pair in uploaded-row mode, a missing FLOW_STAGE column uses the filename as the file identifier. The automatic app also accepts other layouts with a detected `CHCODE`, `primary_key`, or `indexone_uid` header without adding FLOW_STAGE. A successful check shows **MAPPED ROWS MATCH** and covers the automatically mapped fields. If a FLOW_STAGE column is present, its usual nonblank, consistent-value checks still apply. Explicit engine, command-line, and full-copy checks keep their existing stage rules.

The supplied MSB PL CSV was checked against `DataGrid.xlsx`: **640 of 640 rows matched across all five mapped fields**, with no differing values, missing keys, or blank keys. `cc` and `bcc` are blank in this sample and are not included in that result.

The supplied CBS APD CSV was checked against `DataGrid (1).xlsx` (1,105 database rows): **175 of 186 uploaded rows matched across all six mapped fields**. Eleven rows have different email values; both identifiers, names, account numbers, and campaign buckets match for every row. No lookup keys are missing or blank. See `reports/CBS_APD_check_2026-10-22.xlsx` for those differences and the field coverage.

The supplied CBS PL CSV was checked against its refreshed `DataGrid.xlsx` (474 database rows), explicitly using `indexone_uid` → `Old I.C.`: **474 of 474 rows matched across all six mapped fields**. Both identifiers, email, name, account number, and campaign bucket match, with no differing values, missing keys, or blank keys. See `reports/CBS_PL_check_2026-10-22.xlsx` for the results and field coverage. This database is a different snapshot from the earlier MSB PL sample.

## SALAD and PL NEW ENDO email blasts

Upload `SALAD - 10282026.xlsx` as the database and `28-10-26 SBF SALAD EMAIL BLAST NEW ENDO.csv` as the file to check, then select **Check files**. The app recognizes the source workbook's `caseNo`, `caseAccountNo`, `caseCustomerName`, `caseCustomerEmail`, and `Collector` headers. Future exports with the same layout work without renaming columns or adding FLOW_STAGE.

For regular PL endorsements, upload `REGULAR - 10282026.xlsx` as the database and `28-10-26 SBF PL EMAIL BLAST NEW ENDO.csv` as the file to check. This workbook uses the same required source columns and mappings. Extra columns such as `Level_Today` and `PRODUCT_TYPE`, different column positions, and different filenames do not change the account lookup.

| Email-blast column | Endorsement workbook column | Check |
| --- | --- | --- |
| `accountnumber` | `caseAccountNo` | Account lookup; text identifiers preserve leading zeros |
| `chname` | `caseCustomerName` | Customer name on the same account |
| `email` | `caseCustomerEmail` | Email on the same account |

Only these three fields are checked. `agentcode` is **NOT CHECKED** against `Collector`. `cc`, `bcc`, `indexone_uid`, and `primary_key` remain **NOT CHECKED**; blank legacy identifiers do not prevent account-number lookup. The header row is excluded. A present FLOW_STAGE still needs consistent, nonblank values. Existing customer-name comparison rules apply, and original values remain in the report.

Wrong names or emails are **DIFFERENT** even if they belong to another account in the workbook. Unknown and blank account numbers are **NOT FOUND** and **BLANK KEY**. Duplicate accounts require all three mapped values to match one database row. Missing required campaign columns are errors. Other DataGrid and CHCODE layouts keep their existing lookup rules.

The supplied blast contains 16 data rows; all 16 match across the three selected fields. The workbook contains 17 accounts. The remaining account has a blank email and is absent from the blast; campaign subsets are allowed. See `reports/SBF_SALAD_NEW_ENDO_check_2026-10-28.xlsx`.

The supplied PL NEW ENDO blast and REGULAR workbook each contain 10 accounts. All 10 match by account number, name and email, with no blank, missing, or repeated account numbers and no unused workbook accounts. See `reports/SBF_PL_NEW_ENDO_check_2026-10-28.xlsx`.

The command-line checker detects this pair too:

```powershell
.\.venv\Scripts\python.exe scripts\check_files.py --database "..\AUTOCHECKER\SALAD - 10282026.xlsx" --files "..\AUTOCHECKER\28-10-26 SBF SALAD EMAIL BLAST NEW ENDO.csv" --output "reports\SBF_SALAD_NEW_ENDO_check_2026-10-28.xlsx"
```

For the PL pair, use the same command with `REGULAR - 10282026.xlsx`, `28-10-26 SBF PL EMAIL BLAST NEW ENDO.csv`, and `reports/SBF_PL_NEW_ENDO_check_2026-10-28.xlsx`.

## Headerless files and row results

Files without column headers are detected automatically. The app reads column B as CHCODE and includes the first row as data. Only CHCODE values are checked; other columns appear as `Column A`, `Column C`, and so on, and stay **NOT CHECKED** because their meaning cannot be established from a position alone. Headerless files are identified by filename and do not need a FLOW_STAGE column. Successful checks show **MAPPED ROWS MATCH**, covering CHCODE only. The database still needs its original column headers, including its lookup key.

For files with headers, the detected header row and any preamble are excluded from account checks and row totals. Data starts after the header; original source row numbers are retained in results.

Uploaded-row results (the key is CHCODE for campaign files and the detected identifier for DataGrid blasts):

- **MATCHED**: all mapped values match a database row with the same key.
- **DIFFERENT**: the key exists, but one or more mapped values differ.
- **NOT FOUND**: the key does not exist in the database.
- **BLANK KEY**: the uploaded row has no key.

A file receives **EXACT MATCH** only when all its uploaded rows match across the mapped fields and its `FLOW_STAGE` identifier is valid. Unmapped fields remain **NOT CHECKED** and appear in the interface and exported report. A header-only file is an error, rather than a passing check. Repeated uploaded keys are each validated independently against the database.

## FLOW_STAGE identifies the file

`FLOW_STAGE` labels the uploaded file; it is not a field that needs to exist in the database. For example, a file with `FLOW_STAGE=NO_REPLY` is identified as **NO_REPLY**, then each row is checked against its database record using `CHCODE`. The label alone never causes a file to pass.

The supplied CHCODE files use `NO_REPLY`, `NEW_ENDO`, and `SHORT_TERM_ARRANGEMENT`. All three currently use the same row-comparison rule. Where a stage is required or present, every row must have the same nonblank stage. Header letter case is ignored, and stage labels are displayed in uppercase. The automatic app uses filenames for headerless files and files with a detected `CHCODE`, `primary_key`, or `indexone_uid` header when FLOW_STAGE is absent. Explicit engine and command-line checks retain their existing stage requirements; recognized HSBC mobile and DataGrid pairs can omit the stage in uploaded-row mode. Full-copy checks still require a stage.

The result summary and Excel summary show each file's stage. Field coverage marks `FLOW_STAGE` as **FILE IDENTIFIER**. It is excluded from database field mappings and is preserved in the exported source rows. An exact match covers the selected database fields; campaign-only columns without mappings remain unchecked.

## HSBC campaign mappings

Direct header matches are selected automatically. These aliases also support the supplied HSBC campaign layout:

| Uploaded column | Database column | Comparison |
| --- | --- | --- |
| `ChCode` | `CHCODE` | Text |
| `primary_key` | `CHCODE` | Text |
| `indexone_uid` | `CHCODE` | Text |
| `email` | `Email address` | Text |
| `chname` | `Name` | Text |
| `chname_format` | `Last Name` | Text |
| `product_type` | `PROD TYPE` | Product |
| `ProductType` | `PROD TYPE` | Product |
| `LastName` | `Last Name` | Text |
| `formatted mobile` | `Mobile phone` | Phone |
| `OB_AT_ENDORSEMENT` | `OB` | Number |

Field comparisons use:

- **Text**: complete-value equality, ignoring letter case and outer spaces.
- **Number**: numeric equality with commas ignored, so `1,234.50` matches Excel number `1234.5`.
- **Product**: `Credit Card` equals `CARDS`; `Personal Loan` equals `LOAN`. Other product values use normal text comparison.
- **Phone**: equivalent single Philippine mobile formats (`+639`, `639`, `09`, `9`) match. Unrecognized values use text comparison.

Fields without a supported database equivalent are left unchecked: `cc`, `bcc`, `endo`, `COHORT`, `AUDIENCE_SEGMENT`, and `TEMPLATE_FAMILY`. `FLOW_STAGE` is used separately as the file identifier. For example, `endo` contains campaign labels such as `16-18OCT ENDO`; the app does not assume these equal an individual account's referral date.

The supplied samples were checked against `HSBC DB as of Oct 19, 2026.xlsx` (9,277 database rows):

| CSV | Rows | Matched across the seven mapped fields |
| --- | --- | --- |
| 12OCT No Reply | 123 | 123 |
| 16-18OCT New Endo | 320 | 320 |
| 30SEP Short Term Arrangement | 14 | 14 |

The updated sample report is `reports/HSBC_FLOW_STAGE_check_2026-10-19.xlsx`.

## Full database copy comparison

Use the command-line `--full-copy` option when every file should contain the entire database. The Streamlit app always checks uploaded rows against the database. Full-copy mode compares data columns and counts duplicate row occurrences; `FLOW_STAGE` remains a file identifier and is excluded from the database data comparison. Column and row order can differ, but missing or extra data columns, missing rows, additional duplicate rows, or changed values prevent **SAME ROWS**. Field mappings apply only to uploaded-row mode.

Identical rows are paired first. Remaining rows with the same lookup key are paired in source order for cell comparison. Extra occurrences get **EXTRA ROW**; database rows absent from the file get **MISSING FROM FILE**. The report includes a missing-rows sheet for each file in this mode.

## File and matching rules

- Column headers are detected from `CHCODE`, supported identifier headers such as `primary_key` and `indexone_uid`, or supported database layouts; titles, notes, and blank rows above them are skipped. The strict reader requires a name for every data column. Automatic assignment-template and unfamiliar-layout inspection can label populated unnamed columns, such as `Unnamed column K`, without shifting or discarding values. Files without headers use column B as CHCODE and keep row 1. With headers, later repeated names get visible suffixes, such as `CHCode [2]` or `blank [2]`; every column and its values are preserved. The first `CHCODE` column takes priority in ordinary CHCODE layouts. Header comparisons ignore letter case and outer spaces.
- The app selects a single supported worksheet automatically. A workbook with multiple supported worksheets needs to be split into separate files, so no worksheet is silently omitted.
- Text comparisons ignore letter case and trim outer spaces. Product-family aliases are recognized without case differences.
- Text key `00123` remains distinct from `123`. Store keys and account numbers as text when leading zeros matter; zeros shown only by Excel number formatting are not part of the stored value.
- Numeric Excel `123` and `123.0` match text `123`. Lookup keys always use text comparison rules.
- Every nonempty database row needs a value in its lookup key (`CHCODE`, `primary_key`, `indexone_uid`, or DataGrid's `Old I.C.`). Completely empty rows are skipped; source row numbers are preserved.
- When database keys repeat, uploaded-row mode accepts any candidate matching every mapped field. If none match, the closest candidate by number of differing mapped fields is shown.
- File errors do not stop other files. Replacing the uploaded database or files clears old results. A database error must be corrected before checking.
- Excel formula cells use the values last saved by Excel. Recalculate and save workbooks in Excel before checking formulas. A formula without a cached value is treated as blank.
- CSV supports UTF-8 (including BOM) and Windows-1252, with comma, semicolon, tab, or pipe delimiters and standard double-quote escaping. Apostrophes in names are preserved as ordinary text.
- Convert older `.xls` or actual `.accdb` / `.sqlite` databases to Excel or CSV first.
- All rows are loaded into memory. Preview shows up to 500 rows per file; reports contain every checked row. Very large files depend on available memory and Excel's worksheet limits.

Source files are read without modification. Report text beginning with `=` is saved as text rather than an executable formula.

## Command line

```powershell
python scripts/check_files.py --database "database.xlsx" --files "campaign1.csv" "campaign2.csv" --output "reports/check.xlsx"
```

Add `--auto` to use the same detection, inferred mappings, unnamed-column handling, and optional missing `FLOW_STAGE` behavior as the browser app. It saves the browser adapter's generated Excel report and prints the usual per-file JSON summaries. `--auto` cannot be combined with `--database-sheet`, `--database-key`, `--file-key`, or `--full-copy`; use the explicit mode below when those controls are needed. Both modes reject an output path that would overwrite an input file.

Without `--auto`, use `--database-sheet "Sheet1"` to choose a database worksheet, or `--full-copy` to require a complete database copy. This explicit mode reads files with headers; without overrides, lookup columns follow the recognized header layout. Use `--file-key "indexone_uid"` to override the lookup column for every file in that command, and optionally `--database-key "Old I.C."` to override the database key. Header names ignore letter case and outer spaces. A file missing the requested header receives an error while other files continue; it never silently falls back to another key. Run separate commands when files need different lookup columns. The JSON output and Excel summary include the selected lookup columns. Established field mappings use the shared comparison engine; new layout inference is available through `--auto`.

```powershell
python scripts/check_files.py --database "C:\Users\SPM\Desktop\HSBC\DataGrid.xlsx" --files "C:\Users\SPM\Desktop\HSBC\22-10-26 MSB PL EMAIL BLAST.csv" --output "reports/MSB_PL_check.xlsx"
```

```powershell
python scripts/check_files.py --database "C:\Users\SPM\Desktop\HSBC\hiasdh\DataGrid (1).xlsx" --files "C:\Users\SPM\Desktop\HSBC\hiasdh\22-10-26 CBS APD EMAIL BLAST BUCKET 4 91-120.csv" --output "reports/CBS_APD_check.xlsx"
```

```powershell
python scripts/check_files.py --database "C:\Users\SPM\Desktop\HSBC\DataGrid.xlsx" --files "C:\Users\SPM\Desktop\HSBC\22-10-26 CBS PL EMAIL BLAST.csv" --file-key "indexone_uid" --database-key "Old I.C." --output "reports/CBS_PL_check_2026-10-22.xlsx"
```

## Verification

```powershell
python -m unittest discover -s tests -v
```

The suite covers the comparison engine, temporary browser-upload adapter, and Streamlit workflows using Streamlit's AppTest. The file uploader is mocked in AppTest; the actual upload adapter is tested separately with Excel and CSV bytes.

The earlier Windows desktop interface is still available with `python app.py`. Its optional smoke check, `python scripts/check_ui.py`, requires access to the Windows desktop and installed Tcl/Tk libraries. It is not imported by the Streamlit entrypoint and is not needed in Cloud.
