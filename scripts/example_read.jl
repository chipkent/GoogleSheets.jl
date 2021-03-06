
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
