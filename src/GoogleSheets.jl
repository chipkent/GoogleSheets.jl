
#
# Details on the Google Sheets API can be found at:
# https://developers.google.com/sheets/api/guides/concepts
#
# The quickstart guide is also useful
# https://developers.google.com/sheets/api/quickstart/python
#


"""
A package for working with Google Sheets.
"""
module GoogleSheets

using PyCall
using RateLimiter
using JSON
import MacroTools
import DataFrames: DataFrame, nrow, ncol, names
import ColorTypes: Colorant, RGBA, red, green, blue, alpha
using Colors


include("enums.jl")
include("types.jl")
include("json.jl")
include("client.jl")
include("spreadsheet.jl")
include("sheet.jl")
include("meta.jl")
include("io.jl")
include("format.jl")
include("format_conditional.jl")


#TODO: add chart -> https://developers.google.com/sheets/api/reference/rest/v4/spreadsheets/charts
#TODO: add filterview -> batch_update -> https://developers.google.com/sheets/api/guides/filters
#TODO: add basic formatting -> https://developers.google.com/sheets/api/samples/formatting
#TODO: add conditional formatting -> https://developers.google.com/sheets/api/samples/conditional-formatting  https://developers.google.com/sheets/api/guides/conditional-format

end # module
