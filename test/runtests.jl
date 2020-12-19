
using Test, GoogleSheets
using GoogleSheets: CellRangeValues, UpdateSummary

client = sheets_client(AUTH_SPREADSHEET_READWRITE)

spreadsheet_id = "1pG4OyAdePAelCT2fSBTVJ9lVYo6M-ApuTyeEPz49DOM"  #TODO ENV[***]
spreadsheet = Spreadsheet(spreadsheet_id)
sheet = "TestSheet"

function init_test(;add_values::Bool=true)
    try
        delete_sheet!(client, spreadsheet, sheet)
    catch e
    end
    
    add_sheet!(client, spreadsheet, sheet)

    if(add_values)
        # Add values to the sheet
        result = update!(client, CellRange(spreadsheet, sheet), fill(11, 5, 5))
        @test result == UpdateSummary(CellRange(spreadsheet, "$(sheet)!A1:E5"), 5, 5, 25)
        result = get(client, CellRange(spreadsheet, sheet))
        @test result == CellRangeValues(CellRange(spreadsheet, "$(sheet)!A1:Z1000"), fill("11", 5, 5), "ROWS")
    end
end

# Define struct equality.  The default equality uses === for comparisions.
Base.:(==)(x::CellRangeValues, y::CellRangeValues) = x.range == y.range && x.values == y.values && x.major_dimension == y.major_dimension

################################################################################

init_test(add_values=false)

# Get the empty sheet
result = get(client, CellRange(spreadsheet, sheet))
@test result == CellRangeValues(CellRange(spreadsheet, "$(sheet)!A1:Z1000"), nothing, "ROWS")

# Add values to the sheet
result = update!(client, CellRange(spreadsheet, sheet), fill(11, 5, 5))
@test result == UpdateSummary(CellRange(spreadsheet, "$(sheet)!A1:E5"), 5, 5, 25)
result = get(client, CellRange(spreadsheet, sheet))
@test result == CellRangeValues(CellRange(spreadsheet, "$(sheet)!A1:Z1000"), fill("11", 5, 5), "ROWS")

result = get(client, CellRanges(spreadsheet, ["$(sheet)!A1:B2", "$(sheet)!C2:D5"]))
@test result == [
    CellRangeValues(CellRange(spreadsheet, "$(sheet)!A1:B2"), fill("11",2,2), "ROWS"),
    CellRangeValues(CellRange(spreadsheet, "$(sheet)!C2:D5"), fill("11",4,2), "ROWS")
]

################################################################################

init_test()

# Add rows and columns to the sheet
append!(client, spreadsheet, sheet, 1000, 3)
result = get(client, CellRange(spreadsheet, sheet))
@test result == CellRangeValues(CellRange(spreadsheet, "$(sheet)!A1:AC2000"), fill("11", 5, 5), "ROWS")

################################################################################

init_test()

# Insert rows
insert_rows!(client, spreadsheet, sheet, 2, 3, false)
result = get(client, CellRange(spreadsheet, sheet))
values = [fill("11",5), fill("11",5), Any[], fill("11",5), fill("11",5), fill("11",5) ]
@test result == CellRangeValues(CellRange(spreadsheet, "$(sheet)!A1:Z1001"), values, "ROWS")

# Delete rows
delete_rows!(client, spreadsheet, sheet, 2, 3)
result = get(client, CellRange(spreadsheet, sheet))
values = fill("11", 5, 5)
@test result == CellRangeValues(CellRange(spreadsheet, "$(sheet)!A1:Z1000"), values, "ROWS")

################################################################################

init_test()

# Insert columns
insert_cols!(client, spreadsheet, sheet, 2, 3, false)
result = get(client, CellRange(spreadsheet, sheet))
values = fill("11", 5, 6)
values[:,3] .= ""
@test result == CellRangeValues(CellRange(spreadsheet, "$(sheet)!A1:AA1000"), values, "ROWS")

# Delete columns
delete_cols!(client, spreadsheet, sheet, 2, 3)
result = get(client, CellRange(spreadsheet, sheet))
values = fill("11", 5, 5)
@test result == CellRangeValues(CellRange(spreadsheet, "$(sheet)!A1:Z1000"), values, "ROWS")

################################################################################

init_test()

result = clear!(client, CellRange(spreadsheet, "$(sheet)!B2:C3"))
@test result == UpdateSummary(CellRange(spreadsheet, "$(sheet)!B2:C3"), 2, 2, 4)
result = get(client, CellRange(spreadsheet, sheet))
values = fill("11", 5, 5)
values[2:3,2:3] .= ""
@test result == CellRangeValues(CellRange(spreadsheet, "$(sheet)!A1:Z1000"), values, "ROWS")

result = clear!(client, CellRange(spreadsheet, sheet))
@test result == UpdateSummary(CellRange(spreadsheet, "$(sheet)!A1:E5"), 5, 5, 25)
result = get(client, CellRange(spreadsheet, sheet))
@test result == CellRangeValues(CellRange(spreadsheet, "$(sheet)!A1:Z1000"), nothing, "ROWS")

################################################################################

init_test()

freeze!(client, spreadsheet, sheet, 2, 3)

m = meta(client, spreadsheet)
@test spreadsheet_id == m["spreadsheetId"]
show(client, spreadsheet)

m = meta(client, spreadsheet, sheet)
@test sheet == m["title"]
@test 2 == m["gridProperties"]["frozenRowCount"]
@test 3 == m["gridProperties"]["frozenColumnCount"]
show(client, spreadsheet, sheet)

sheet_id = m["sheetId"]
m = meta(client, spreadsheet, sheet_id)
@test sheet == m["title"]
@test 2 == m["gridProperties"]["frozenRowCount"]
@test 3 == m["gridProperties"]["frozenColumnCount"]
show(client, spreadsheet, sheet_id)

################################################################################
