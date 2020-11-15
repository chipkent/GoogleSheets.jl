
using GoogleSheets

# Example based upon: # https://developers.google.com/sheets/api/quickstart/python

auth = auth_service()

# The ID and range of a sample spreadsheet.
SAMPLE_SPREADSHEET_ID = "1BxiMVs0XRA5nFMdKvBdBZjgmUUqptlbs74OgvE2upms"
SAMPLE_RANGE_NAME = "Class Data!A2:E"

sheet = Spreadsheet(SAMPLE_SPREADSHEET_ID)
range = CellRange(sheet, SAMPLE_RANGE_NAME)
result = get(auth, range)

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

    println("")
    println("Name, Major:")
    for row in eachrow(values)
        # Print columns A and E, which correspond to indices 1 and 5.
        println("ROW: $(row[1]), $(row[5])")
    end
end
