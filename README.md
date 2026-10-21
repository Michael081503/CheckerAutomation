# CHCODE Database Checker

A Streamlit app for validating rows in multiple Excel or CSV files against a database workbook. `FLOW_STAGE` identifies files with headers; headerless files use their filenames. `CHCODE` locates each corresponding database row. The app compares mapped fields and provides an Excel report download. It is prepared for Streamlit Community Cloud.

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
2. Add multiple files in **Add files to check**. Files with headers need one consistent, nonblank `FLOW_STAGE` identifier. For headerless files, select the column B option described below.
3. Click **Check files**. The default settings compare your uploaded rows against the database; smaller campaign subsets are supported.
4. Read the compact results, then expand a file to see its **Issues**, **Rows**, or **Fields checked**. Row previews default to entries needing review when issues exist.
5. Click **Download Excel report** for every original row, changed value, and field mapping.

Optional settings stay under **Advanced options**: choose worksheets, change the comparison mode, or use **Edit comparison fields** to adjust mappings. Files with the same column layout and header setting share a mapping table. CHCODE lookup is always retained; files with headers also retain FLOW_STAGE identification.

The database and uploaded files can have titles, notes, or blank rows above their column headers. The app automatically uses the first row containing a `CHCODE` header, ignoring letter case and outer spaces, and skips the rows above it. It displays the detected database header row and lists each uploaded file's header row under **Advanced options**. Leave **No headers (CHCODE in column B)** unchecked when a `CHCODE` header exists below the first row. Source row numbers in results and reports retain their original positions.

For a file without column headers, open **Advanced options** and select **No headers (CHCODE in column B)** for that file. The app reads column B as CHCODE and includes the first row as data. It checks CHCODE values against the database by default; other columns appear as `Column A`, `Column C`, and so on, and stay **NOT CHECKED** until you map them in **Edit comparison fields**. Headerless files are identified by filename and do not need a FLOW_STAGE column. Successful checks show **MAPPED ROWS MATCH**, covering only the selected fields. This option supports uploaded-row validation; the database still needs its original column headers, including CHCODE.

Uploaded-row results:

- **MATCHED**: all mapped values match a database row with the same CHCODE.
- **DIFFERENT**: the CHCODE exists, but one or more mapped values differ.
- **NOT FOUND**: the CHCODE does not exist in the database.
- **BLANK KEY**: the uploaded row has no CHCODE.

A file receives **EXACT MATCH** only when all its uploaded rows match across the selected mapped fields and its `FLOW_STAGE` identifier is valid. Unmapped fields remain **NOT CHECKED** and appear in the interface and exported report. A header-only file is an error, rather than a passing check. Repeated uploaded CHCODEs are each validated independently against the database.

## FLOW_STAGE identifies the file

`FLOW_STAGE` labels the uploaded file; it is not a field that needs to exist in the database. For example, a file with `FLOW_STAGE=NO_REPLY` is identified as **NO_REPLY**, then each row is checked against its database record using `CHCODE`. The label alone never causes a file to pass.

The supplied files use `NO_REPLY`, `NEW_ENDO`, and `SHORT_TERM_ARRANGEMENT`. All three currently use the same row-comparison rule. In files with headers, every row must have the same stage; missing headers, blank stages, or mixed stages produce a file error. Header letter case is ignored, and stage labels are displayed in uppercase. Headerless uploads use their filenames as identifiers instead.

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

Identical rows are paired first. Remaining rows with the same CHCODE are paired in source order for cell comparison. Extra occurrences get **EXTRA ROW**; database rows absent from the file get **MISSING FROM FILE**. The report includes a missing-rows sheet for each file in this mode.

## File and matching rules

- Column headers are detected on the first row containing a `CHCODE` header; titles, notes, and blank rows above it are skipped. Each data column needs a name on that header row. Select **No headers (CHCODE in column B)** only for uploaded files without headers. With headers, later repeated names get visible suffixes, such as `CHCode [2]` or `blank [2]`; every column and its values are preserved. The first `CHCODE` column is the default key. Header comparisons ignore letter case and outer spaces.
- Ignore letter case and trim outer spaces are enabled by default for text. Disable them for case-sensitive values and significant spaces. Product-family aliases are recognized without case differences.
- Text CHCODE `00123` remains distinct from `123`. Store codes as text when leading zeros matter; zeros shown only by Excel number formatting are not part of the stored value.
- Numeric Excel `123` and `123.0` match text `123`. CHCODEs always use text comparison rules.
- Every nonempty database row needs a CHCODE. Completely empty rows are skipped; source row numbers are preserved.
- When database CHCODEs repeat, uploaded-row mode accepts any candidate matching every mapped field. If none match, the closest candidate by number of differing mapped fields is shown.
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

Use `--database-sheet "Sheet1"` to choose a database worksheet, or `--full-copy` to require a complete database copy. Otherwise this command uses the same suggested field mappings as the Streamlit app. Custom field mappings are selected in the Streamlit app.

## Verification

```powershell
python -m unittest discover -s tests -v
```

The suite covers the comparison engine, temporary browser-upload adapter, and Streamlit workflows using Streamlit's AppTest. The file uploader is mocked in AppTest; the actual upload adapter is tested separately with Excel and CSV bytes.

The earlier Windows desktop interface is still available with `python app.py`. Its optional smoke check, `python scripts/check_ui.py`, requires access to the Windows desktop and installed Tcl/Tk libraries. It is not imported by the Streamlit entrypoint and is not needed in Cloud.
