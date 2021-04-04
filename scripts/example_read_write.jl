
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
