
using GoogleSheets

# Example based upon: # https://developers.google.com/sheets/api/quickstart/python

client = sheets_client(AUTH_SPREADSHEET_READWRITE)

# The ID and range of a sample spreadsheet.
SAMPLE_SPREADSHEET_ID = "1pG4OyAdePAelCT2fSBTVJ9lVYo6M-ApuTyeEPz49DOM"
SAMPLE_RANGE_NAME = "Sheet1"

sheet = Spreadsheet(SAMPLE_SPREADSHEET_ID)
range = CellRange(sheet, SAMPLE_RANGE_NAME)
result = get(client, range)

println("KEYS: $(keys(result))")
println("RANGE: $(result["range"])")
println("MAJORDIM: $(result["majorDimension"])")

# values = result["values"]
#
# if isnothing(values)
#     println("No data found.")
# else
#     for row in eachrow(values)
#         println("ROW: $row")
#     end
#
#     println("")
#     println("Name, Major:")
#     for row in eachrow(values)
#         # Print columns A and E, which correspond to indices 1 and 5.
#         println("ROW: $(row[1]), $(row[5])")
#     end
# end

values = ["0" "1" "2"; "a" "=A1+B1" 33]
println(values)

result = update!(client, range, values)

println()
show(client, sheet)

try
    delete_sheet!(client, sheet, "test sheet")
    println("Deleted sheet")
catch e
    println("No sheet to delete")
end

add_sheet!(client, sheet, "test sheet")

println()
show(client, sheet, "test sheet")
