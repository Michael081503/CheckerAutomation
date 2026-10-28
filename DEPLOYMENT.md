# Deploy the CHCODE checker to Streamlit Community Cloud

The Cloud entrypoint is **`streamlit_app.py`**. The app uses the existing CHCODE comparison engine and works with browser uploads; it does not depend on Windows paths or Tkinter.

## Files to put in GitHub

Keep this layout at the root of the repository:

```text
streamlit_app.py
web_checker.py
checker.py
requirements.txt
.gitignore
README.md
DEPLOYMENT.md
.streamlit/
    config.toml
```

`artifacts/streamlit-cloud-ready.zip`, if provided, contains this clean source layout. Extract it before adding the files to GitHub; Streamlit deploys from the repository, not directly from the ZIP. Local desktop files and tests are optional for deployment.

Upload actual database and campaign files through the app. They are not required in GitHub. The `.gitignore` excludes Excel and CSV data, generated reports, temporary upload folders, and `.streamlit/secrets.toml`.

## Deploy

1. Create a GitHub repository and add the source files above. Include `.streamlit/config.toml` so the dark theme and upload setting are applied.
2. Open [Streamlit Community Cloud](https://share.streamlit.io/) and select **Create app**.
3. Select your repository and its branch, normally `main`.
4. Set **Main file path** to **`streamlit_app.py`**.
5. In **Advanced settings**, choose **Python 3.13**, matching the version used for local verification. No secrets or system packages are required by this app.
6. Select **Deploy**. Community Cloud installs `requirements.txt` and starts the app.

These repository, entrypoint, and Python-version settings follow Streamlit's [official deployment guide](https://docs.streamlit.io/deploy/streamlit-community-cloud/deploy-your-app/deploy). The pinned Streamlit version is `1.54.0`; openpyxl is `3.1.5`.

## Recover after a source update

The entrypoint refreshes an older in-memory checker and upload adapter before using the current comparison rules. This handles a Cloud process that reruns the updated app while retaining modules from the previous release. Normal reruns preserve the loaded modules and current results. A comparison-version check also refreshes releases that already supported Debtor ID but used older customer-name rules.

If Cloud still reports `cannot import name 'is_datagrid_pl_viber_blast'`, verify the app uses this repository's `main` branch and `streamlit_app.py`, then select **Manage app → ⋮ → Reboot app** and confirm **Reboot**. This clears the process memory and reloads the repository source. See Streamlit's [reboot instructions](https://docs.streamlit.io/deploy/streamlit-community-cloud/manage-your-app/reboot-your-app).

## Use the deployed app

1. Upload the database under **Add your database**.
2. Upload multiple CSV or Excel files under **Add files to check**. Files with and without headers can be mixed. For the MSB or CBS PL email blast, upload the corresponding `DataGrid.xlsx` snapshot as the database; for the CBS APD email blast, upload `DataGrid (1).xlsx`.
3. Review the automatic detection summary: it shows each file's worksheet, header row, lookup column, and status. The database must contain `CHCODE`, `primary_key`, or `indexone_uid`, the recognized DataGrid headers (`Name`, `Old I.C.`, `Account No.`, and `Email`), or the SBF endorsement headers described below. A workbook with multiple supported worksheets must be split into separate files before upload.
4. Select **Check files**. The app compares uploaded rows against the database and automatically maps the supported HSBC or DataGrid fields. Campaign subsets are supported.
5. Read the compact pass/review results. Expand an individual file for its issues, row preview, or field coverage, then download the Excel report.

Headers, lookup columns, worksheets, and supported field mappings are detected automatically. There are no manual setup controls. Uploaded headers named `CHCODE`, `primary_key`, or `indexone_uid` can contain CHCODE, at any column position; keep the original header names. `FLOW_STAGE` identifies the file type, such as `NO_REPLY`, and is not compared to a database column. The database does not need a `FLOW_STAGE` header. Files with a detected lookup header and files without headers use their filenames when a stage is absent. A present stage must contain one consistent, nonblank value. **EXACT MATCH** means every uploaded row matches its database record across the mapped fields and its stage is valid. Files checked without a stage show **MAPPED ROWS MATCH** when all mapped values match. Other unmapped campaign fields are listed in the expanded field coverage and report.

DataGrid email blasts match their detected identifier to `Old I.C.` and compare both identifiers, email, name, and account number when present. If both `primary_key` and `indexone_uid` exist, the app uses the column with more values matching the uploaded database across all rows. Identical columns with matching keys use `primary_key`, including the SBF layout. Other ties or files with no matching keys retain the layout's usual choice: `primary_key` when `campaignbucket` is present, otherwise `indexone_uid` when available. A `CHCODE` header takes priority. `campaignbucket` is also compared with `Account Type` when that database column exists, giving both supplied CBS files six checked fields. Matching results show **MAPPED ROWS MATCH**; differing emails appear under **Issues** and in the report. Fields without a supported counterpart, such as `cc` and `bcc`, remain unchecked.

For `23-10-26 SBF PL EMAIL BLAST.csv`, the app detects `primary_key` as CHCODE. It also compares `financingnumber` with `Financing/Card No.` when that database field exists, giving six checked fields: both identifiers, email, name, account number, and financing number.

Files without headers use column B as CHCODE and include the first row as data. Only CHCODE is checked for these files; the other columns remain **NOT CHECKED**. Original row numbers and values are preserved in the report. The database still needs its column headers.

For the HSBC mobile blasts, add `HSBC DB as of Oct 23, 2026.xlsx` as the database and the Field and NEW ENDO workbooks as files to check. When FLOW_STAGE is absent, each `formatted mobile` is compared with `Mobile phone` for the same CHCODE. All five supported fields remain checked: phone, CHCODE, name, last name, and product type. Successful checks show **MAPPED ROWS MATCH**. Wrong or blank phone numbers show **DIFFERENT**, with expected values from the same account in the app and report. Equivalent Philippine mobile formats are normalized. Files with phone headers need a CHCODE lookup column; the database also needs a supported phone column. A present stage receives the usual validation and mapped account comparison.

For the SBF SALAD Viber blasts, add `DataGrid salad.xlsx` as the database and the Level 1, 3, and 5 CSVs as files to check. The app recognizes their full ten-column layout and matches `Phone` with `Last Called No.`, then checks `FullName` against `Name` on the same DataGrid row. Both fields must match to pass. The report lists unknown phones, blank phones, and changed names, along with the actual database lookup column. Other campaign fields, including `CHCode` and campaign amounts, remain **NOT CHECKED**. This format can omit FLOW_STAGE; other CHCODE layouts retain their account-based checks.

Question marks are ignored when comparing customer names in both uploaded and database rows. Original names remain in source rows and reports, and genuine name differences still require review.

For SALAD NEW ENDO email blasts, upload the `SALAD - 10282026.xlsx` endorsement workbook and the matching CSV. Its accountnumber is looked up in caseAccountNo; chname and email are compared with caseCustomerName and caseCustomerEmail on that same row. Agentcode is not compared with Collector, and blank legacy identifier columns remain unchecked. The source workbook is recognized by its caseNo, caseAccountNo, caseCustomerName, caseCustomerEmail, and Collector headers. Header rows are excluded, and campaign subsets are allowed.

For PL NEW ENDO email blasts, upload `REGULAR - 10282026.xlsx` and `28-10-26 SBF PL EMAIL BLAST NEW ENDO.csv`. The same account-number lookup and three field checks apply. Extra regular-workbook columns, including Level_Today and PRODUCT_TYPE, do not affect detection. Staff assignments remain unchecked.

In customer names, the replacement character `�` matches a corresponding `Ñ` or `ñ` at the same position on the same account. Every other character and mapped field must agree. It is not a wildcard for other letters, IDs, emails, or notes. Files with headers exclude the header row and preamble from account checks and totals; files without headers retain their first data row.

Replacing an uploaded database or file clears old results. Run the check again after changes. Files with errors are reported while other files continue. A database error must be corrected before a check can run.

The report includes all checked rows, individual cell differences, and field coverage. Browser previews show up to 500 rows at a time. Full database copy comparison remains available locally through the command-line `--full-copy` option. Explicit engine and command-line checks keep their existing FLOW_STAGE rules, including the full-copy stage requirement; the automatic app allows an absent stage for every detected lookup header. See [README.md](README.md#command-line).

## Data lifecycle and limits

- Uploads are transferred to the Streamlit server and kept in that browser session. Inspection metadata, comparison results, and report bytes use session state rather than a global cache.
- Each check writes inputs into an isolated temporary directory and removes it when the check finishes or fails. Reports use uploaded filenames rather than server temporary paths.
- **Start over** resets the current app session and shows empty upload controls. Refreshing or disconnecting can also reset Streamlit session state; save your downloaded report if you need to keep it.
- The app has no database, external API calls, permanent upload storage, or configured login. App access is managed through your Community Cloud sharing settings.
- `.streamlit/config.toml` sets a 200 MB limit per uploaded file, matching Streamlit's default. Uploaded workbooks expand in memory during parsing, so available Cloud resources can limit practical file sizes before that upload limit.

For upload-limit configuration, see the official [`st.file_uploader` documentation](https://docs.streamlit.io/develop/api-reference/widgets/st.file_uploader). Change `server.maxUploadSize` in `.streamlit/config.toml` if your environment needs a different limit.

## Run locally

```powershell
python -m pip install -r requirements.txt
python -m streamlit run streamlit_app.py
```

On Windows, double-click **Start Checker.bat** after installing dependencies. Its terminal must stay open while the app runs.

## Verify

```powershell
python -m unittest discover -s tests -v
```

Tests cover CHCODE matching, automatic header and lookup detection, HSBC aliases, numeric/product conversions, duplicate columns, temporary-file cleanup, simultaneous filenames, upload errors, Streamlit previews, report generation, upload replacement, and session reset. They use synthetic data. Actual file-picker uploads should also be tried once in the deployed browser app.
