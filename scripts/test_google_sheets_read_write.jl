
using GoogleSheets

# Example based upon: # https://developers.google.com/sheets/api/quickstart/python

auth = auth_service(AUTH_SPREADSHEET_READWRITE)

# The ID and range of a sample spreadsheet.
SAMPLE_SPREADSHEET_ID = "1pG4OyAdePAelCT2fSBTVJ9lVYo6M-ApuTyeEPz49DOM"
SAMPLE_RANGE_NAME = "Sheet1"

sheet = Spreadsheet(SAMPLE_SPREADSHEET_ID)
range = CellRange(sheet, SAMPLE_RANGE_NAME)
result = get(auth, range)

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

try
    global result = update(auth, range, values)
catch e
    println("ERROR: $e")
    println("STACK: $(e.traceback)")
    using PyCall
    tb = pyimport("traceback")
    tb.print_tb(e.traceback)
    tb.print_exception(e.traceback)
end
