export DataFrame, get, update!, clear!


"""
Creates a DataFrame from spreadsheet range values.  The first row is converted to column names.  All other rows are converted to string values.
"""
DataFrame(values::CellRangeValues)::Union{Nothing,DataFrame} = values.values == nothing ? nothing : DataFrame([values.values[1,i]=>values.values[2:end,i] for i in 1:size(values.values,2)]...)
 

"""
Converts sheet values to a string matrix.
"""
function _matrix(inputs::Array{String,2})::Array{String,2}
    return inputs
end


"""
Converts sheet values to a string matrix.
"""
function _matrix(inputs)::Array{String,2}
    m = fill("", length(inputs), mapreduce(length, max, inputs))

    for (i,r) in enumerate(inputs)
        for (j,v) in enumerate(r)
            m[i,j] = v
        end
    end

    return m
end


"""
Gets a range of cell values from a spreadsheet.
"""
function Base.get(client::GoogleSheetsClient, range::CellRange)::CellRangeValues
    result = gsheet_api_sheet_get(client; spreadsheetId=range.spreadsheet.id, majorDimension="ROWS", range=range.range)
    return CellRangeValues(CellRange(range.spreadsheet, result["range"]), haskey(result, "values") ? _matrix(result["values"]) : nothing, result["majorDimension"])
end


"""
Gets multiple ranges of cell values from a spreadsheet.
"""
function Base.get(client::GoogleSheetsClient, ranges::CellRanges)::Vector{CellRangeValues}
    result = gsheet_api_sheet_batchget(client; spreadsheetId=ranges.spreadsheet.id, majorDimension="ROWS", ranges=ranges.ranges)
    return [CellRangeValues(CellRange(ranges.spreadsheet, r["range"]), haskey(r, "values") ? _matrix(r["values"]) : nothing, r["majorDimension"]) for r in result["valueRanges"] ]
end


"""
Updates a range of cell values in a spreadsheet.

# Arguments
- `client::GoogleSheetsClient`: client for interacting with Google Sheets.
- `range::CellRange`: cell range to modify.
- `values::Array{<:Any,2}`: values to place in the spreadsheet.
- `raw::Bool=false`: true treats values as raw, unparsed values and and are simply
    inserted as a string.  false treats values exactly as if they were entered into
    the Google Sheets UI, for example "=A1+B1" is a formula.
"""
function update!(client::GoogleSheetsClient, range::CellRange, values::Array{<:Any,2}; raw::Bool=false)::UpdateSummary
    # There are serialization problems if the values are not Array{Any,2} or Array{String,2}
    values = Any[values[i,j] for i in 1:size(values,1), j in 1:size(values,2)]

    body = Dict(
        "values" => values,
        "majorDimension" => "ROWS",
    )

    result = gsheet_api_sheet_update(client; spreadsheetId=range.spreadsheet.id, range=range.range, valueInputOption= raw ? "RAW" : "USER_ENTERED", body=body)
    return UpdateSummary(CellRange(range.spreadsheet, result["updatedRange"]), result["updatedColumns"], result["updatedRows"], result["updatedCells"])
end


"""
Updates a range of cell values in a spreadsheet.

# Arguments
- `client::GoogleSheetsClient`: client for interacting with Google Sheets.
- `range::CellRange`: cell range to modify.
- `df::DataFrame`: dataframe of values.
- `kwargs...`: keyword arguments.
"""
function update!(client::GoogleSheetsClient, range::CellRange, df::DataFrame; kwargs...)::UpdateSummary

    function df_to_string(df::DataFrame)::Matrix{String}
        rst = fill("", nrow(df)+1, ncol(df))
    
        for (j,n) in enumerate(names(df))
            rst[1, j] = n
        end
    
        for i in 1:nrow(df), j in 1:ncol(df)
            rst[i+1,j] = "$(df[i,j])"
        end
    
        return rst
    end

    return update!(client, range, df_to_string(df); kwargs...)
end


"""
Clears a range of cell values in a spreadsheet.
"""
function clear!(client::GoogleSheetsClient, range::CellRange)::UpdateSummary
    v = get(client, range)
    rng = v.range
    vls = v.values

    if isnothing(values)
        throw(ErrorException("No data found: range=$range"))
    end

    vls .= ""
    return update!(client, CellRange(range.spreadsheet, rng.range), vls)
end

