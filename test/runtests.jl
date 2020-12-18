
using Test, GoogleSheets

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
        result = get(client, CellRange(spreadsheet, sheet))
        @test result == Dict{Any,Any}("range" => "$(sheet)!A1:Z1000", "values" => fill("11", 5, 5), "majorDimension" => "ROWS")
    end
end

################################################################################

init_test(add_values=false)

# Get the empty sheet
result = get(client, CellRange(spreadsheet, sheet))
@test !haskey(result, "values")

# Add values to the sheet
result = update!(client, CellRange(spreadsheet, sheet), fill(11, 5, 5))
result = get(client, CellRange(spreadsheet, sheet))
@test result == Dict{Any,Any}("range" => "$(sheet)!A1:Z1000", "values" => fill("11", 5, 5), "majorDimension" => "ROWS")

result = get(client, CellRanges(spreadsheet, ["$(sheet)!A1:B2", "$(sheet)!C2:D5"]))
@test result == Dict{Any,Any}("spreadsheetId" => spreadsheet_id, "valueRanges" => Dict{Any,Any}[
    Dict("range" => "$(sheet)!A1:B2", "values" => fill("11",2,2), "majorDimension" => "ROWS"),
    Dict("range" => "$(sheet)!C2:D5", "values" => fill("11",4,2), "majorDimension" => "ROWS"),
])

################################################################################

init_test()

# Add rows and columns to the sheet
append!(client, spreadsheet, sheet, 1000, 3)
result = get(client, CellRange(spreadsheet, sheet))
@test result == Dict{Any,Any}("range" => "$(sheet)!A1:AC2000", "values" => fill("11", 5, 5), "majorDimension" => "ROWS")

################################################################################

init_test()

# Insert rows
insert_rows!(client, spreadsheet, sheet, 2, 3, false)
result = get(client, CellRange(spreadsheet, sheet))
values = [fill("11",5), fill("11",5), Any[], fill("11",5), fill("11",5), fill("11",5) ]
@test result == Dict{Any,Any}("range" => "$(sheet)!A1:Z1001", "values" => values, "majorDimension" => "ROWS")

# Delete rows
delete_rows!(client, spreadsheet, sheet, 2, 3)
result = get(client, CellRange(spreadsheet, sheet))
values = fill("11", 5, 5)
@test result == Dict{Any,Any}("range" => "$(sheet)!A1:Z1000", "values" => values, "majorDimension" => "ROWS")

################################################################################

init_test()

# Insert columns
insert_cols!(client, spreadsheet, sheet, 2, 3, false)
result = get(client, CellRange(spreadsheet, sheet))
values = fill("11", 5, 6)
values[:,3] .= ""
@test result == Dict{Any,Any}("range" => "$(sheet)!A1:AA1000", "values" => values, "majorDimension" => "ROWS")

# Delete columns
delete_cols!(client, spreadsheet, sheet, 2, 3)
result = get(client, CellRange(spreadsheet, sheet))
values = fill("11", 5, 5)
@test result == Dict{Any,Any}("range" => "$(sheet)!A1:Z1000", "values" => values, "majorDimension" => "ROWS")

################################################################################

init_test()

clear!(client, CellRange(spreadsheet, "$(sheet)!B2:C3"))
result = get(client, CellRange(spreadsheet, sheet))
values = fill("11", 5, 5)
values[2:3,2:3] .= ""
@test result == Dict{Any,Any}("range" => "$(sheet)!A1:Z1000", "values" => values, "majorDimension" => "ROWS")

clear!(client, CellRange(spreadsheet, sheet))
result = get(client, CellRange(spreadsheet, sheet))
@test result == Dict{Any,Any}("range" => "$(sheet)!A1:Z1000", "majorDimension" => "ROWS")

################################################################################

init_test()

# just make sure these don't exception
freeze!(client, spreadsheet, sheet, 2, 3)

m = meta(client, spreadsheet)
@test spreadsheet_id == m["spreadsheetId"]
show(client, spreadsheet)

m = meta(client, spreadsheet, sheet)
@test sheet == m["title"]
show(client, spreadsheet, sheet)

sheet_id = m["sheetId"]
m = meta(client, spreadsheet, sheet_id)
@test sheet == m["title"]
show(client, spreadsheet, sheet_id)

################################################################################
