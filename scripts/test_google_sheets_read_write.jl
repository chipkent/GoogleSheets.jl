
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

println("KEYS: $(keys(result))")
println("RANGE: $(result["range"])")
println("MAJORDIM: $(result["majorDimension"])")

values = result["values"]

if isnothing(values)
    println("No data found.")
else
    for row in eachrow(values)
        println("ROW: $row")
    end
end

################################################################################

result = get(client, ranges)

println("KEYS: $(keys(result))")
println("RANGE: $(result["valueRanges"])")

values = result["valueRanges"]

if isnothing(values)
    println("No data found.")
else
    for (k,v) in values
        for row in eachrow(v[2])
            println("ROW: $(k[2]) $row")
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

freeze!(client, sheet, "test sheet", 2, 3)
append!(client, sheet, "test sheet", 1000, 3)

insert_rows!(client, sheet, "test sheet", 2, 3, false)
insert_cols!(client, sheet, "test sheet", 2, 3, false)

delete_rows!(client, sheet, "test sheet", 2, 3)
delete_cols!(client, sheet, "test sheet", 2, 3)

clear!(client, CellRange(sheet, "test sheet!B2:C3"))
