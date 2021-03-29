
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
using JSON
import MacroTools
import DataFrames: DataFrame, nrow, ncol, names
import ColorTypes: Colorant, RGBA, red, green, blue, alpha
using Colors


"""
Print details on a python exception.
"""
macro _print_python_exception(ex)
    # MacroTools.@q is used instead of quote so that the returned stacktrace
    # has line numbers from the calling function and not the macro.
    return esc(MacroTools.@q begin
        try
            $ex
        catch e
            if hasfield(typeof(e), :traceback)
                println("Python error:")
                println(e)
                println("Python stacktrace:")
                tb = pyimport("traceback")
                tb.print_exception(e.traceback)
                tb.print_tb(e.traceback)
            end
            rethrow(e)
        end
    end)
end


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


#TODO: ******** next work is to figure out how to handle sheet names / ids as well as the ranges.

#TODO: cell range stuff is used in sheet.

#TODO: rename to gsheet_json and move to json.jl
"""
Returns a dictionary of values to describe a 1D range of cells in a sheet.
"""
function cellRange1D(sheet::Sheet, dim::AbstractString, start_index::Integer, end_index::Integer)
    return Dict(
        "sheetId" => sheet.id,
        "dimension" => dim,
        "startIndex" => start_index,
        "endIndex" => end_index,
    )
end


#TODO: rename to gsheet_json and move to json.jl
"""
Returns a dictionary of values to describe a 2D range of cells in a sheet.
"""
function cellRange2D(sheet::Sheet, start_row_index::Integer, end_row_index::Integer, start_col_index::Integer, end_col_index::Integer)
    return Dict(
            "sheetId" => sheet.id,
            "startRowIndex" => start_row_index,
            "endRowIndex" => end_row_index+1,
            "startColumnIndex" => start_col_index,
            "endColumnIndex" => end_col_index+1,
        )
end


#TODO: add chart -> https://developers.google.com/sheets/api/reference/rest/v4/spreadsheets/charts
#TODO: add filterview -> batch_update -> https://developers.google.com/sheets/api/guides/filters
#TODO: add basic formatting -> https://developers.google.com/sheets/api/samples/formatting
#TODO: add conditional formatting -> https://developers.google.com/sheets/api/samples/conditional-formatting  https://developers.google.com/sheets/api/guides/conditional-format

end # module
