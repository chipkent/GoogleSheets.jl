
using Test, GoogleSheets
using GoogleSheets: CellRangeValues, UpdateSummary
using ColorTypes
using Colors

# #TODO: remove this 
cf = GoogleSheets.credentials_file()
println("CONFIGFILE: ", cf)
println("ISFILE: ", isfile(cf))
println("FILE: ", read(cf, String))

if !haskey(ENV, "SPREADSHEET_ID")
    error("The environment variable SPREADSHEET_ID is not defined")
end

client = sheets_client(AUTH_SCOPE_READWRITE)

spreadsheet_id = ENV["SPREADSHEET_ID"]
spreadsheet = Spreadsheet(spreadsheet_id)
sheet_name = "TestSheet"

function init_test(;add_values::Bool=true)
    if sheet_name in sheet_names(client,spreadsheet)
        delete_sheet!(client, spreadsheet, sheet_name)
    end
    
    add_sheet!(client, spreadsheet, sheet_name)
    @test ["Sheet1", "TestSheet"] == sheet_names(client,spreadsheet)
    @test [Sheet(client,spreadsheet,"Sheet1"),Sheet(client,spreadsheet,"TestSheet")] == sheets(client,spreadsheet)

    if(add_values)
        # Add values to the sheet
        result = update!(client, CellRange(spreadsheet, sheet_name), fill(11, 5, 5))
        @test result == UpdateSummary(CellRange(spreadsheet, "$(sheet_name)!A1:E5"), 5, 5, 25)
        result = get(client, CellRange(spreadsheet, sheet_name))
        @test result == CellRangeValues(CellRange(spreadsheet, "$(sheet_name)!A1:Z1000"), fill("11", 5, 5), "ROWS")
    end
end

# Define struct equality.  The default equality uses === for comparisions.
Base.:(==)(x::CellRangeValues, y::CellRangeValues) = x.range == y.range && x.values == y.values && x.major_dimension == y.major_dimension

################################################################################

values = ["A" "B" "C"; "1" "2" "3"; "4" "5" "6"]
crv = CellRangeValues(CellRange(spreadsheet, "$(sheet_name)!A1:"), values, "TEST_DIM")
df = DataFrame(crv)
@test df == DataFrame(A= ["1", "4"], B=["2","5"], C=["3","6"])

################################################################################

init_test(add_values=false)

# Get the empty sheet
result = get(client, CellRange(spreadsheet, sheet_name))
@test result == CellRangeValues(CellRange(spreadsheet, "$(sheet_name)!A1:Z1000"), nothing, "ROWS")

# Add values to the sheet
df = DataFrame(A=[1,2], B=["X","Y"])
result = update!(client, CellRange(spreadsheet, sheet_name), df)
@test result == UpdateSummary(CellRange(spreadsheet, "$(sheet_name)!A1:B3"), 2, 3, 6)
result = get(client, CellRange(spreadsheet, sheet_name))
@test result == CellRangeValues(CellRange(spreadsheet, "$(sheet_name)!A1:Z1000"), ["A" "B"; "1" "X"; "2" "Y"], "ROWS")

################################################################################

init_test(add_values=false)

# Get the empty sheet
result = get(client, CellRange(spreadsheet, sheet_name))
@test result == CellRangeValues(CellRange(spreadsheet, "$(sheet_name)!A1:Z1000"), nothing, "ROWS")

# Add values to the sheet
result = update!(client, CellRange(spreadsheet, sheet_name), fill(11, 5, 5))
@test result == UpdateSummary(CellRange(spreadsheet, "$(sheet_name)!A1:E5"), 5, 5, 25)
result = get(client, CellRange(spreadsheet, sheet_name))
@test result == CellRangeValues(CellRange(spreadsheet, "$(sheet_name)!A1:Z1000"), fill("11", 5, 5), "ROWS")

result = get(client, CellRanges(spreadsheet, ["$(sheet_name)!A1:B2", "$(sheet_name)!C2:D5"]))
@test result == [
    CellRangeValues(CellRange(spreadsheet, "$(sheet_name)!A1:B2"), fill("11",2,2), "ROWS"),
    CellRangeValues(CellRange(spreadsheet, "$(sheet_name)!C2:D5"), fill("11",4,2), "ROWS")
]

################################################################################

init_test()

# Add rows and columns to the sheet
sheet = Sheet(client, spreadsheet, sheet_name)
append!(client, sheet, 1000, 3)
result = get(client, CellRange(spreadsheet, sheet_name))
@test result == CellRangeValues(CellRange(spreadsheet, "$(sheet_name)!A1:AC2000"), fill("11", 5, 5), "ROWS")

################################################################################

init_test()

# Insert rows
sheet = Sheet(client, spreadsheet, sheet_name)
insert_rows!(client, CellIndexRange1D(sheet, 2, 3))
result = get(client, CellRange(spreadsheet, sheet_name))
values = fill("11", 6, 5)
values[3,:] .= ""
@test result == CellRangeValues(CellRange(spreadsheet, "$(sheet_name)!A1:Z1001"), values, "ROWS")

# Delete rows
sheet = Sheet(client, spreadsheet, sheet_name)
delete_rows!(client, CellIndexRange1D(sheet, 2, 3))
result = get(client, CellRange(spreadsheet, sheet_name))
values = fill("11", 5, 5)
@test result == CellRangeValues(CellRange(spreadsheet, "$(sheet_name)!A1:Z1000"), values, "ROWS")

################################################################################

init_test()

# Insert columns
sheet = Sheet(client, spreadsheet, sheet_name)
insert_cols!(client, CellIndexRange1D(sheet, 2, 3))
result = get(client, CellRange(spreadsheet, sheet_name))
values = fill("11", 5, 6)
values[:,3] .= ""
@test result == CellRangeValues(CellRange(spreadsheet, "$(sheet_name)!A1:AA1000"), values, "ROWS")

# Delete columns
sheet = Sheet(client, spreadsheet, sheet_name)
delete_cols!(client, CellIndexRange1D(sheet, 2, 3))
result = get(client, CellRange(spreadsheet, sheet_name))
values = fill("11", 5, 5)
@test result == CellRangeValues(CellRange(spreadsheet, "$(sheet_name)!A1:Z1000"), values, "ROWS")

################################################################################

init_test()

result = clear!(client, CellRange(spreadsheet, "$(sheet_name)!B2:C3"))
@test result == UpdateSummary(CellRange(spreadsheet, "$(sheet_name)!B2:C3"), 2, 2, 4)
result = get(client, CellRange(spreadsheet, sheet_name))
values = fill("11", 5, 5)
values[2:3,2:3] .= ""
@test result == CellRangeValues(CellRange(spreadsheet, "$(sheet_name)!A1:Z1000"), values, "ROWS")

result = clear!(client, CellRange(spreadsheet, sheet_name))
@test result == UpdateSummary(CellRange(spreadsheet, "$(sheet_name)!A1:E5"), 5, 5, 25)
result = get(client, CellRange(spreadsheet, sheet_name))
@test result == CellRangeValues(CellRange(spreadsheet, "$(sheet_name)!A1:Z1000"), nothing, "ROWS")

################################################################################

init_test()

sheet = Sheet(client, spreadsheet, sheet_name)
freeze!(client, sheet, 2, 3)

m = meta(client, spreadsheet)
@test spreadsheet_id == m["spreadsheetId"]
show(client, spreadsheet)

m = meta(client, spreadsheet, sheet_name)
@test sheet_name == m["title"]
@test 2 == m["gridProperties"]["frozenRowCount"]
@test 3 == m["gridProperties"]["frozenColumnCount"]
show(client, spreadsheet, sheet_name)

sheet_id = m["sheetId"]
m = meta(client, spreadsheet, sheet_id)
@test sheet_name == m["title"]
@test 2 == m["gridProperties"]["frozenRowCount"]
@test 3 == m["gridProperties"]["frozenColumnCount"]
show(client, spreadsheet, sheet_id)

m = meta(client, sheet)
@test sheet_name == m["title"]
@test 2 == m["gridProperties"]["frozenRowCount"]
@test 3 == m["gridProperties"]["frozenColumnCount"]
show(client, sheet)

################################################################################

init_test()

sheet = Sheet(client, spreadsheet, sheet_name)

cf = CellFormat(
    background_color=RGBA(0.5, 0.5, 0.5, 0.8), number_format_type=NUMBER_FORMAT_TYPE_NUMBER, number_format_pattern="0.0", 
    text_italic=true, text_bold=true, text_strikethrough=true, text_color=RGB(0.5, 0.5, 0.5), text_font_size=14)

cf2 = CellFormat(background_color=RGBA(0.5, 0.5, 0.5, 0.8), text_italic=true, text_bold=true, text_strikethrough=true)
    
format!(client, CellIndexRange2D(sheet, 2, 3, 1, 2), CellFormat(number_format_type=NUMBER_FORMAT_TYPE_NUMBER, number_format_pattern="0.0"))
format!(client, CellIndexRange2D(sheet, 4, 5, 1, 2), CellFormat(number_format_type=NUMBER_FORMAT_TYPE_DATE, number_format_pattern="hh:mm:ss am/pm, ddd mmm dd yyyy"))
format!(client, CellIndexRange2D(sheet, 4, 5, 1, 2), CellFormat(background_color=RGBA(0.5, 0.5, 0.5, 0.8)))
format!(client, CellIndexRange2D(sheet, 4, 5, 1, 2), CellFormat(background_color=RGB(0.5, 0.5, 0.5)))
format!(client, CellIndexRange2D(sheet, 4, 5, 1, 2), CellFormat(background_color=Gray(0.5)))
format!(client, CellIndexRange2D(sheet, 2, 3, 1, 2), cf)

format_conditional!(client, CellIndexRange2D(sheet, 2, 3, 1, 2), cf2, CONDITION_TYPE_NUMBER_GREATER_THAN_EQ, 0.5)
format_conditional!(client, CellIndexRange2D(sheet, 2, 3, 1, 2), cf2, CONDITION_TYPE_NUMBER_BETWEEN, 0.5, 1.5)

format_color_gradient!(client, CellIndexRange2D(sheet, 4, 5, 1, 2))
format_color_gradient!(client, CellIndexRange2D(sheet, 4, 5, 1, 2); min_value_type=VALUE_TYPE_NUMBER, min_value=-3, max_value_type=VALUE_TYPE_NUMBER, max_value=3, mid_color=colorant"white", mid_value_type=VALUE_TYPE_NUMBER, mid_value=0)

################################################################################
