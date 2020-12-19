# GoogleSheets.jl

Julia package for working with Google Sheets.

Key functions:
* sheets_client
* meta
* show
* add_sheet!
* delete_sheet!
* insert_rows!
* delete_rows!
* insert_cols!,
* delete_cols!
* append!
* get
* update!
* clear!
* batch_update!
* freeze!

To use:
1. Create a Google Sheets API token from either the [python quick start reference](https://developers.google.com/sheets/api/quickstart/python) or the [developers console](https://console.developers.google.com/apis/credentials).
2. Place the Google Sheets API `credentials.json` file in `~/.julia/google_sheets/`.
3. Connect to Google Sheets using `sheets_client`.
4. See the scripts directory for examples of using the package.
