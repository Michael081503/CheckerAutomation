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

## Use the deployed app

1. Upload the database under **Add your database**.
2. Upload multiple CSV or Excel files under **Add files to check**. Each uploaded file needs one consistent, nonblank `FLOW_STAGE` identifier.
3. Select **Check files**. Default settings support campaign subsets and automatically map the seven HSBC fields.
4. Read the compact pass/review results. Expand an individual file for its issues, row preview, or field coverage, then download the Excel report.

**Advanced options** contains the optional comparison settings, worksheets, and **Edit comparison fields** controls. `FLOW_STAGE` identifies the file type, such as `NO_REPLY`, and is not compared to a database column. The database does not need a `FLOW_STAGE` header. **EXACT MATCH** means every uploaded row matches its database record across the mapped fields; a stage label alone does not establish a match. Other unmapped campaign fields are listed in the expanded field coverage and report.

Upload, worksheet, mapping, or comparison-option changes invalidate old results. Run the check again after changes. Files with errors are reported while other files continue. A database error must be corrected before a check can run.

The report includes all checked rows, individual cell differences, and field coverage. Full database copy mode instead includes missing database rows. Browser previews show up to 500 rows at a time.

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

Tests cover CHCODE matching, HSBC aliases, numeric/product conversions, duplicate columns, temporary-file cleanup, simultaneous filenames, upload errors, Streamlit previews, report generation, option changes, upload replacement, and session reset. They use synthetic data. Actual file-picker uploads should also be tried once in the deployed browser app.
