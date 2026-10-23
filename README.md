# CHCODE Database Checker

A Streamlit app for validating rows in multiple Excel or CSV files against a database workbook. It supports CHCODE campaign files and DataGrid email blasts. `CHCODE` locates campaign database rows; recognized DataGrid blasts use `indexone_uid` or `primary_key` to look up `Old I.C.`. Choose the lookup column for each file when the same headers represent different campaign layouts. The app compares mapped fields and provides an Excel report download. It is prepared for Streamlit Community Cloud.

The interface uses a dark theme with compact upload panels, a teal check button, and results that expand when you need details. Advanced settings stay collapsed by default.

## Start

From this folder, install the pinned dependencies and launch Streamlit:

```powershell
python -m pip install -r requirements.txt
python -m streamlit run streamlit_app.py
```

On Windows, **Start Checker.bat** launches the same app in your browser. Keep the terminal open while using the local app. The Streamlit app requires Python 3.10+; it was tested on Python 3.13.

## Deploy to Streamlit Community Cloud

Upload the app source to a GitHub repository, then select that repository in Streamlit Community Cloud with **Main file path: `streamlit_app.py`** and **Python 3.13** in Advanced settings. See [DEPLOYMENT.md](DEPLOYMENT.md) for the exact file list and steps. The official [deployment guide](https://docs.streamlit.io/deploy/streamlit-community-cloud/deploy-your-app/deploy) describes the repository and entrypoint settings.

Upload your databases and campaign files through the app after deployment. The repository's `.gitignore` excludes Excel/CSV inputs, generated reports, uploads, and Streamlit secrets.

Uploaded data is processed on the app server. Each browser session keeps its own inspection metadata, results, and downloadable report. The app does not use a shared data cache or save uploads to the repository. Temporary files are deleted after inspection or checking, including failures. **Start over** resets the current app session. Uploads and results do not survive a server restart.

## Check your files

1. Add the database `.xlsx`, `.xlsm`, or `.csv` file in **Add your database**.
2. Add multiple files in **Add files to check**. CHCODE files with headers need one consistent, nonblank `FLOW_STAGE` identifier. Recognized DataGrid email blasts can use their filenames instead. For headerless files, select the column B option described below.
3. Click **Check files**. The default settings compare your uploaded rows against the database; smaller campaign subsets are supported.
4. Read the compact results, then expand a file to see its **Issues**, **Rows**, or **Fields checked**. Row previews default to entries needing review when issues exist.
5. Click **Download Excel report** for every original row, changed value, and field mapping.

Optional settings stay under **Advanced options**: choose worksheets and lookup columns, change the comparison mode, or use **Edit comparison fields** to adjust mappings. **Database lookup column (CHCODE)** selects the database key; each file's **Lookup column (CHCODE)** selects its key. Recognized DataGrid databases select `Old I.C.` automatically. Files with the same column layout, header setting, and lookup column share a mapping table. The lookup-key mapping is always retained. CHCODE files with headers also retain FLOW_STAGE identification. Changing a lookup column clears the previous results and report.

The database and uploaded files can have titles, notes, or blank rows above their column headers. The app detects `CHCODE` or the recognized DataGrid header layout, ignoring letter case and outer spaces, and skips the rows above it. It displays the detected database header row and lists each uploaded file's header row under **Advanced options**. Leave **No headers (CHCODE in column B)** unchecked when the file has headers, including DataGrid files. Source row numbers in results and reports retain their original positions.

## DataGrid email blasts

To check `22-10-26 MSB PL EMAIL BLAST.csv`, add `DataGrid.xlsx` as the database, add the CSV as a file to check, and click **Check files** with the default uploaded-row mode. Keep **No headers (CHCODE in column B)** unchecked: both files already have headers. No CHCODE or FLOW_STAGE columns need to be added.

To check `22-10-26 CBS APD EMAIL BLAST BUCKET 4 91-120.csv`, add `DataGrid (1).xlsx` as the database and the CBS APD CSV as a file to check. Click **Check files** with the same defaults. The checker uses the CBS file's `primary_key` column as its CHCODE lookup and compares it with the database's `Old I.C.`.

To check `22-10-26 CBS PL EMAIL BLAST.csv`, add its `DataGrid.xlsx` database and the CBS PL CSV. Open **Advanced options**, leave **Database lookup column (CHCODE)** at `Old I.C.`, and select `indexone_uid` under **Lookup column (CHCODE): 22-10-26 CBS PL EMAIL BLAST.csv**. Then click **Check files**. This explicitly uses `indexone_uid` as CHCODE against `Old I.C.`.

The checker recognizes a database with `Name`, `Old I.C.`, `Account No.`, and `Email`, paired with a blast containing `email`, `chname`, `accountnumber`, and `indexone_uid` or `primary_key`. Automatic selection uses `primary_key` when both `campaignbucket` and `primary_key` exist, including the CBS APD sample. Other recognized blasts default to `indexone_uid` when available, otherwise `primary_key`. A `CHCODE` header takes priority by default. CBS PL and CBS APD can share identical headers, so select the file lookup column that contains CHCODE for that campaign. Suggested comparisons are:

| Email-blast column | DataGrid column | Comparison |
| --- | --- | --- |
| `indexone_uid` | `Old I.C.` | Text |
| `primary_key` | `Old I.C.` | Text |
| `email` | `Email` | Text |
| `chname` | `Name` | Text |
| `accountnumber` | `Account No.` | Text |
| `campaignbucket` | `Account Type` | Text, when both columns exist |

Both identifier columns are checked when present, including an `indexone_uid` that differs from the selected `primary_key`. Leading zeros in text identifiers and account numbers are preserved. `cc` and `bcc` remain **NOT CHECKED** unless you select a database mapping.

For this recognized pair in uploaded-row mode, a missing FLOW_STAGE column uses the filename as the file identifier. A successful check shows **MAPPED ROWS MATCH** and covers the selected mapped fields. If a FLOW_STAGE column is present, its usual nonblank, consistent-value checks still apply. Other files with headers keep the existing FLOW_STAGE requirement; full-copy mode keeps it too.

The supplied MSB PL CSV was checked against `DataGrid.xlsx`: **640 of 640 rows matched across all five mapped fields**, with no differing values, missing keys, or blank keys. `cc` and `bcc` are blank in this sample and are not included in that result.

The supplied CBS APD CSV was checked against `DataGrid (1).xlsx` (1,105 database rows): **175 of 186 uploaded rows matched across all six mapped fields**. Eleven rows have different email values; both identifiers, names, account numbers, and campaign buckets match for every row. No lookup keys are missing or blank. See `reports/CBS_APD_check_2026-10-22.xlsx` for those differences and the field coverage.

The supplied CBS PL CSV was checked against its refreshed `DataGrid.xlsx` (474 database rows), explicitly using `indexone_uid` → `Old I.C.`: **474 of 474 rows matched across all six mapped fields**. Both identifiers, email, name, account number, and campaign bucket match, with no differing values, missing keys, or blank keys. See `reports/CBS_PL_check_2026-10-22.xlsx` for the results and field coverage. This database is a different snapshot from the earlier MSB PL sample.

## Headerless files and row results

For a file without column headers, open **Advanced options** and select **No headers (CHCODE in column B)** for that file. The app reads column B as CHCODE and includes the first row as data. It checks CHCODE values against the database by default; other columns appear as `Column A`, `Column C`, and so on, and stay **NOT CHECKED** until you map them in **Edit comparison fields**. Headerless files are identified by filename and do not need a FLOW_STAGE column. Successful checks show **MAPPED ROWS MATCH**, covering only the selected fields. This option supports uploaded-row validation; the database still needs its original column headers, including its lookup key.

Uploaded-row results (the key is CHCODE for campaign files and the selected UID for DataGrid blasts):

- **MATCHED**: all mapped values match a database row with the same key.
- **DIFFERENT**: the key exists, but one or more mapped values differ.
- **NOT FOUND**: the key does not exist in the database.
- **BLANK KEY**: the uploaded row has no key.

A file receives **EXACT MATCH** only when all its uploaded rows match across the selected mapped fields and its `FLOW_STAGE` identifier is valid. Unmapped fields remain **NOT CHECKED** and appear in the interface and exported report. A header-only file is an error, rather than a passing check. Repeated uploaded keys are each validated independently against the database.

## FLOW_STAGE identifies the file

`FLOW_STAGE` labels the uploaded file; it is not a field that needs to exist in the database. For example, a file with `FLOW_STAGE=NO_REPLY` is identified as **NO_REPLY**, then each row is checked against its database record using `CHCODE`. The label alone never causes a file to pass.

The supplied CHCODE files use `NO_REPLY`, `NEW_ENDO`, and `SHORT_TERM_ARRANGEMENT`. All three currently use the same row-comparison rule. Every row must have the same stage; missing headers, blank stages, or mixed stages produce a file error. Header letter case is ignored, and stage labels are displayed in uppercase. Headerless uploads and recognized DataGrid blasts without a FLOW_STAGE column use their filenames as identifiers instead.

The result summary and Excel summary show each file's stage. Field coverage marks `FLOW_STAGE` as **FILE IDENTIFIER**. It is excluded from database field mappings and is preserved in the exported source rows. An exact match covers the selected database fields; campaign-only columns without mappings remain unchecked.

## HSBC campaign mappings

Direct header matches are selected automatically. These aliases also support the supplied HSBC campaign layout:

| Uploaded column | Database column | Comparison |
| --- | --- | --- |
| `ChCode` | `CHCODE` | Text |
| `primary_key` | `CHCODE` | Text |
| `email` | `Email address` | Text |
| `chname` | `Name` | Text |
| `chname_format` | `Last Name` | Text |
| `product_type` | `PROD TYPE` | Product |
| `OB_AT_ENDORSEMENT` | `OB` | Number |

The mapping table supports:

- **Text**: complete-value equality using the case and space options.
- **Number**: numeric equality with commas ignored, so `1,234.50` matches Excel number `1234.5`.
- **Product**: `Credit Card` equals `CARDS`; `Personal Loan` equals `LOAN`. Other product values use normal text comparison.

Fields without a direct database equivalent are left unchecked by default: `cc`, `bcc`, `endo`, `COHORT`, `AUDIENCE_SEGMENT`, and `TEMPLATE_FAMILY`. `FLOW_STAGE` is used separately as the file identifier. For example, `endo` contains campaign labels such as `16-18OCT ENDO`; the app does not assume these equal an individual account's referral date. You can select a database mapping if your workflow defines one.

The supplied samples were checked against `HSBC DB as of Oct 19, 2026.xlsx` (9,277 database rows):

| CSV | Rows | Matched across the seven mapped fields |
| --- | --- | --- |
| 12OCT No Reply | 123 | 123 |
| 16-18OCT New Endo | 320 | 320 |
| 30SEP Short Term Arrangement | 14 | 14 |

The updated sample report is `reports/HSBC_FLOW_STAGE_check_2026-10-19.xlsx`.

## Full database copy comparison

Select **Compare complete database copy** when every file should contain the entire database. This mode compares data columns and counts duplicate row occurrences; `FLOW_STAGE` remains a file identifier and is excluded from the database data comparison. Column and row order can differ, but missing or extra data columns, missing rows, additional duplicate rows, or changed values prevent **SAME ROWS**. Field mappings apply only to uploaded-row mode.

Identical rows are paired first. Remaining rows with the same lookup key are paired in source order for cell comparison. Extra occurrences get **EXTRA ROW**; database rows absent from the file get **MISSING FROM FILE**. The report includes a missing-rows sheet for each file in this mode.

## File and matching rules

- Column headers are detected from `CHCODE` or the supported DataGrid layout; titles, notes, and blank rows above them are skipped. Each data column needs a name on that header row. Select **No headers (CHCODE in column B)** only for uploaded files without headers. With headers, later repeated names get visible suffixes, such as `CHCode [2]` or `blank [2]`; every column and its values are preserved. The first `CHCODE` column takes priority as the default key when present. Header comparisons ignore letter case and outer spaces.
- Ignore letter case and trim outer spaces are enabled by default for text. Disable them for case-sensitive values and significant spaces. Product-family aliases are recognized without case differences.
- Text key `00123` remains distinct from `123`. Store keys and account numbers as text when leading zeros matter; zeros shown only by Excel number formatting are not part of the stored value.
- Numeric Excel `123` and `123.0` match text `123`. Lookup keys always use text comparison rules.
- Every nonempty database row needs a value in its lookup key (`CHCODE` or DataGrid's `Old I.C.`). Completely empty rows are skipped; source row numbers are preserved.
- When database keys repeat, uploaded-row mode accepts any candidate matching every mapped field. If none match, the closest candidate by number of differing mapped fields is shown.
- File errors do not stop other files. Changes to the database, files, worksheets, mappings, or check options clear old results. Inputs are read again when checking; avoid editing them during a check.
- Excel formula cells use the values last saved by Excel. Recalculate and save workbooks in Excel before checking formulas. A formula without a cached value is treated as blank.
- CSV supports UTF-8 (including BOM) and Windows-1252, with comma, semicolon, tab, or pipe delimiters.
- Convert older `.xls` or actual `.accdb` / `.sqlite` databases to Excel or CSV first.
- All rows are loaded into memory. Preview shows up to 500 rows per file; reports contain every checked row. Very large files depend on available memory and Excel's worksheet limits.

Source files are read without modification. Report text beginning with `=` is saved as text rather than an executable formula.

## Command line

```powershell
python scripts/check_files.py --database "database.xlsx" --files "campaign1.csv" "campaign2.csv" --output "reports/check.xlsx"
```

Use `--database-sheet "Sheet1"` to choose a database worksheet, or `--full-copy` to require a complete database copy. Otherwise this command detects CHCODE or the supported DataGrid layout and uses the same suggested field mappings as the Streamlit app. Use `--file-key "indexone_uid"` to override the lookup column for every file in that command, and optionally `--database-key "Old I.C."` to override the database key. Header names ignore letter case and outer spaces. A file missing the requested header receives an error while other files continue; it never silently falls back to another key. Run separate commands when files need different lookup columns. The JSON output and Excel summary include the selected lookup columns. Custom field mappings are selected in the Streamlit app.

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
