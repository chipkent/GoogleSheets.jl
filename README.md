<!-- ![Test](https://github.com/chipkent/GoogleSheets.jl/actions/workflows/test.yml/badge.svg) -->
![Register](https://github.com/chipkent/GoogleSheets.jl/actions/workflows/register.yml/badge.svg)
![Document](https://github.com/chipkent/GoogleSheets.jl/actions/workflows/document.yml/badge.svg)
![Compat Helper](https://github.com/chipkent/GoogleSheets.jl/actions/workflows/compathelper.yml/badge.svg)
![Tagbot](https://github.com/chipkent/GoogleSheets.jl/actions/workflows/tagbot.yml/badge.svg)

# GoogleSheets.jl

Julia package for working with Google Sheets.  You can perform expected actions such as adding sheets,
removing sheets, reading from sheets, writing to sheets, and formatting sheets.

Key types:
* `GoogleSheetsClient`
* `Spreadsheet`
* `Sheet`
* `CellRange`
* `CellRanges`
* `CellRangeValues`
* `UpdateSummary`
* `CellIndexRange1D`
* `CellIndexRange2D`
* `CellFormat`
* `DataFrame`

Key functions:
* `sheets_client`
* `sheet_names`
* `sheets`
* `batch_update!`
* `add_sheet!`
* `delete_sheet!`
* `freeze!`
* `append!`
* `insert_rows!`
* `insert_cols!`
* `delete_rows!`
* `delete_cols!`
* `meta`
* `show`
* `get`
* `update!`
* `clear!`
* `format!`
* `format_conditional!`
* `format_color_gradient!`

To use:
1. Create a Google Sheets API token from either the [python quick start reference](https://developers.google.com/sheets/api/quickstart/python) or the [developers console](https://console.developers.google.com/apis/credentials).
2. Place the Google Sheets API `credentials.json` file in `~/.julia/google_sheets/`.
3. Connect to Google Sheets using `sheets_client`.
4. See the scripts directory for examples of using the package.

An example reading data from a Google Sheet.  See `./scripts/example_read.jl`.

```julia
using GoogleSheets

# Example based upon: # https://developers.google.com/sheets/api/quickstart/python

client = sheets_client(AUTH_SCOPE_READONLY)

# The ID and range of a sample spreadsheet.
SAMPLE_SPREADSHEET_ID = "1BxiMVs0XRA5nFMdKvBdBZjgmUUqptlbs74OgvE2upms"
SAMPLE_RANGE_NAME = "Class Data!A2:E"

sheet = Spreadsheet(SAMPLE_SPREADSHEET_ID)
range = CellRange(sheet, SAMPLE_RANGE_NAME)
result = get(client, range)
println("RESULT: $(result)")

if isnothing(result.values)
    println("No data found.")
else
    for row in eachrow(result.values)
        println("ROW: $row")
    end

    println("")
    println("Name, Major:")
    for row in eachrow(result.values)
        # Print columns A and E, which correspond to indices 1 and 5.
        println("ROW: $(row[1]), $(row[5])")
    end
end
```

An example reading data, writing data, and modifying a Google Sheet.  See `./scripts/example_read_write.jl`.

```julia
using GoogleSheets

# Example based upon: # https://developers.google.com/sheets/api/quickstart/python

client = sheets_client(AUTH_SCOPE_READWRITE)

# The ID and range of a sample spreadsheet.
SAMPLE_SPREADSHEET_ID = "1pG4OyAdePAelCT2fSBTVJ9lVYo6M-ApuTyeEPz49DOM"
SAMPLE_RANGE_NAME = "Sheet1"

sheet = Spreadsheet(SAMPLE_SPREADSHEET_ID)
range = CellRange(sheet, SAMPLE_RANGE_NAME)
ranges = CellRanges(sheet, ["Sheet1!A1:B9", "Sheet1!B1:B9"])

println()
show(client, sheet)

values = ["0" "1" "2"; "a" "=A1+B1" 33]
println(values)

result = update!(client, range, values)


################################################################################

result = get(client, range)
println("RESULT: $(result)")

if isnothing(result.values)
    println("No data found.")
else
    for row in eachrow(result.values)
        println("ROW: $row")
    end
end

################################################################################

result = get(client, ranges)
println("RESULT: $(result)")

for r in result
    if isnothing(r.values)
        println("No data found.")
    else
        for row in eachrow(r.values)
            println("ROW: $(r.range) $row")
        end
    end
end

################################################################################

try
    delete_sheet!(client, sheet, "test sheet")
    println("Deleted sheet")
catch e
    println("No sheet to delete")
end

add_sheet!(client, sheet, "test sheet")

println()
show(client, sheet, "test sheet")

values = fill(11, 5, 5)
println("VALUES $(typeof(values)) $values")
result = update!(client, CellRange(sheet, "test sheet"), values)

freeze!(client, Sheet(client, sheet, "test sheet"), 2, 3)
append!(client, Sheet(client, sheet, "test sheet"), 1000, 3)

insert_rows!(client, CellIndexRange1D(Sheet(client, sheet, "test sheet"), 2, 3))
insert_cols!(client, CellIndexRange1D(Sheet(client, sheet, "test sheet"), 2, 3))

delete_rows!(client, CellIndexRange1D(Sheet(client, sheet, "test sheet"), 2, 3))
delete_cols!(client, CellIndexRange1D(Sheet(client, sheet, "test sheet"), 2, 3))

clear!(client, CellRange(sheet, "test sheet!B2:C3"))
```

## Documentation

See [https://chipkent.github.io/GoogleSheets.jl/dev/](https://chipkent.github.io/GoogleSheets.jl/dev/).

Pull requests will publish documentation to <code>https://chipkent.github.io/GoogleSheets.jl/previews/PR##</code>.