
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

export GoogleSheetsClient, Spreadsheet, CellRange, CellRanges, DataFrame, sheets_client, meta, show, sheet_names, get, update!,
        clear!, batch_update!, add_sheet!, delete_sheet!, freeze!, append!, insert_rows!, insert_cols!,
        delete_rows!, delete_cols!, format_number!, format_datetime!, format_background_color!


"""
Directory containing configuration files.
"""
config_dir = joinpath(homedir(),".julia/config/google_sheets/")


"""
Create an enum and export the enum type and all values.
"""
macro _exported_enum(name, args...)
    return esc(quote
        @enum($name, $(args...))
        export $name
        $([:(export $arg) for arg in args]...)
    end)
end


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


"""
An authorization scope for accessing Google resources.

AUTH_SPREADSHEET_READONLY:  Allows read-only access to the user's sheets and their properties.
AUTH_SPREADSHEET_READWRITE: Allows read/write access to the user's sheets and their properties.
"""
@_exported_enum AuthScope AUTH_SPREADSHEET_READONLY AUTH_SPREADSHEET_READWRITE


const _permission_urls = Dict(
    AUTH_SPREADSHEET_READONLY => "https://www.googleapis.com/auth/spreadsheets.readonly",
    AUTH_SPREADSHEET_READWRITE => "https://www.googleapis.com/auth/spreadsheets",
)


"""
A Google Sheets client.
"""
struct GoogleSheetsClient
    client
end


"""
A spreadsheet.
"""
struct Spreadsheet
    """Spreadsheet unique identifier."""
    id::AbstractString
end


"""
A range of cells within a spreadsheet.
"""
struct CellRange
    """Spreadsheet containing the cells."""
    spreadsheet::Spreadsheet

    """Range of cells."""
    range::AbstractString
end


"""
Multiple ranges of cells within a spreadsheet.
"""
struct CellRanges{T<:AbstractString}
    """Spreadsheet containing the cells."""
    spreadsheet::Spreadsheet

    """Ranges of cells."""
    ranges::Array{T,1}
end


"""
A range of cell values within a spreadsheet.
"""
struct CellRangeValues
    """Range of cells within a spreadsheet."""
    range::CellRange

    """Values of cells within a spreadsheet."""
    values::Union{Nothing,Array{String,2}}

    """Major dimension of the cell values."""
    major_dimension::AbstractString
end


"""
Creates a DataFrame from spreadsheet range values.  The first row is converted to column names.  All other rows are converted to string values.
"""
DataFrame(values::CellRangeValues)::Union{Nothing,DataFrame} = values.values == nothing ? nothing : DataFrame([values.values[1,i]=>values.values[2:end,i] for i in 1:size(values.values,2)]...)
 

"""
Summary of updated updated cells.
"""
struct UpdateSummary
    """Range of cells within a spreadsheet."""
    range::CellRange

    """Number of updated columns."""
    updated_columns::Int64

    """Number of updated rows."""
    updated_rows::Int64

    """Number of updated cells."""
    updated_cells::Int64
end


"""
Maps authorization scopes to the appropriate permission URLs.
"""
function _scope_urls(scopes::AuthScope)::Array{String,1}
    return [_permission_urls[scopes]]
end


"""
Maps authorization scopes to the appropriate permission URLs.
"""
function _scope_urls(scopes::Array{AuthScope,1})::Array{String,1}
    return [_permission_urls[scope] for scope in scopes]
end


"""
Gets the credentials file needed to log into Google.
The file is loaded from the GOOGLE_SHEETS_CREDENTIALS environment variable
if it is present; otherwise, it is loaded from the configuration directory,
which defaults to ~/.julia/config/google_sheets/.

See the python quick start reference for a link to generate credentials.
https://developers.google.com/sheets/api/quickstart/python

See the API credentials page to create or revoke credentials.
https://console.developers.google.com/apis/credentials
"""
function credentials_file()::String
    if haskey(ENV, "GOOGLE_SHEETS_CREDENTIALS")
        return ENV["GOOGLE_SHEETS_CREDENTIALS"]
    end

    file = "credentials.json"
    return joinpath(config_dir, file)
end


"""
Gets a cached token file which allows access to Google Sheets with
specific authorization scopes.
"""
function _token_file(scopes::AuthScope)::String
    return _token_file([scopes])
end


"""
Gets a cached token file which allows access to Google Sheets with
specific authorization scopes.
"""
function _token_file(scopes::Array{AuthScope,1})::String
    s = sort(unique(copy(scopes)))

    id = 0

    for i in 1:length(s)
        id += Int(s[i]) * 10^i
    end

    file = "google_sheets_token.$id.pickle"
    return joinpath(config_dir, file)
end


"""
Creates a client for accessing Google Sheets.
"""
function sheets_client(scopes::Union{AuthScope,Array{AuthScope,1}})::GoogleSheetsClient
    pickle = pyimport("pickle")
    os_path = pyimport("os.path")
    build = pyimport("googleapiclient.discovery").build
    InstalledAppFlow = pyimport("google_auth_oauthlib.flow").InstalledAppFlow
    Request = pyimport("google.auth.transport.requests").Request
    open = pybuiltin("open")

    credentialsFile = credentials_file()
    tokenFile = _token_file(scopes)
    scopeUrls = _scope_urls(scopes)

    @_print_python_exception begin
        creds = nothing

        # The file token.pickle stores the user's access and refresh tokens, and is
        # created automatically when the authorization flow completes for the first
        # time.
        if os_path.exists(tokenFile)
            @pywith open(tokenFile, "rb") as token begin
                creds = pickle.load(token)
            end
        end

        # If there are no (valid) credentials available, let the user log in.
        if isnothing(creds) || !creds.valid
            if !isnothing(creds) && creds.expired && !isnothing(creds.refresh_token)
                creds.refresh(Request())
            else
                flow = InstalledAppFlow.from_client_secrets_file(credentialsFile, scopeUrls)
                creds = flow.run_local_server(port=0)
            end

            # Save the credentials for the next run
            @pywith open(tokenFile, "wb") as token begin
                pickle.dump(creds, token)
            end
        end

        return GoogleSheetsClient(build("sheets", "v4", credentials=creds))
    end
end


"""
Gets metadata about a spreadsheet.
"""
function meta(client::GoogleSheetsClient, spreadsheet::Spreadsheet)::Dict{Any,Any}
    @_print_python_exception begin
        sheet = client.client.spreadsheets()
        result = sheet.get(spreadsheetId=spreadsheet.id).execute()
        return result
    end
end


"""
Gets metadata about a spreadsheet sheet.
"""
function meta(client::GoogleSheetsClient, spreadsheet::Spreadsheet, title::AbstractString)::Dict{Any,Any}
    metadata = meta(client, spreadsheet)
    sheets = metadata["sheets"]

    for sheet in sheets
        properties = sheet["properties"]

        if properties["title"] == title
            return properties
        end
    end

    throw(KeyError(title))
end


"""
Gets metadata about a spreadsheet sheet.
"""
function meta(client::GoogleSheetsClient, spreadsheet::Spreadsheet, sheet_id::Int64)::Dict{Any,Any}
    metadata = meta(client, spreadsheet)
    sheets = metadata["sheets"]

    for sheet in sheets
        properties = sheet["properties"]

        if properties["sheetId"] == sheet_id
            return properties
        end
    end

    throw(KeyError(sheet_id))
end


"""
Prints metadata about a spreadsheet.
"""
function Base.show(client::GoogleSheetsClient, spreadsheet::Spreadsheet)
    m = meta(client, spreadsheet)
    println("Spreadsheet:")
    println(json(m, 4))
end


"""
Prints metadata about a spreadsheet sheet.
"""
function Base.show(client::GoogleSheetsClient, spreadsheet::Spreadsheet, title::AbstractString)
    m = meta(client, spreadsheet, title)
    println("Sheet:")
    println(json(m, 4))
end


"""
Prints metadata about a spreadsheet sheet.
"""
function Base.show(client::GoogleSheetsClient, spreadsheet::Spreadsheet, sheet_id::Int64)
    m = meta(client, spreadsheet, sheet_id)
    println("Sheet:")
    println(json(m, 4))
end


"""
Gets the names of the sheets in the spreadsheet.
"""
function sheet_names(client::GoogleSheetsClient, spreadsheet::Spreadsheet)::Vector{String}
    m = meta(client, spreadsheet)
    return [ s["properties"]["title"] for s in m["sheets"] ]
end


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
    @_print_python_exception begin
        sheet = client.client.spreadsheets()
        result = sheet.values().get(spreadsheetId=range.spreadsheet.id,
                                    majorDimension="ROWS",
                                    range=range.range).execute()
        return CellRangeValues(CellRange(range.spreadsheet, result["range"]), haskey(result, "values") ? _matrix(result["values"]) : nothing, result["majorDimension"])
    end
end


"""
Gets multiple ranges of cell values from a spreadsheet.
"""
function Base.get(client::GoogleSheetsClient, ranges::CellRanges)::Vector{CellRangeValues}
    @_print_python_exception begin
        sheet = client.client.spreadsheets()
        result = sheet.values().batchGet(spreadsheetId=ranges.spreadsheet.id,
                                    majorDimension="ROWS",
                                    ranges=ranges.ranges).execute()

        return [CellRangeValues(CellRange(ranges.spreadsheet, r["range"]), haskey(r, "values") ? _matrix(r["values"]) : nothing, r["majorDimension"]) for r in result["valueRanges"] ]
    end
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

    @_print_python_exception begin
        sheet = client.client.spreadsheets()
        result = sheet.values().update(spreadsheetId=range.spreadsheet.id,
                                range=range.range,
                                valueInputOption= raw ? "RAW" : "USER_ENTERED",
                                body=body).execute()

        return UpdateSummary(CellRange(range.spreadsheet, result["updatedRange"]), result["updatedColumns"], result["updatedRows"], result["updatedCells"])
    end
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


"""
Applies one or more updates to a spreadsheet.

Each request is validated before being applied. If any request is not valid then
the entire request will fail and nothing will be applied.

Common batch_update! functionality:
Charts: https://developers.google.com/sheets/api/reference/rest/v4/spreadsheets/charts
Filters: https://developers.google.com/sheets/api/guides/filters
Basic formatting: https://developers.google.com/sheets/api/samples/formatting
Conditional formatting: https://developers.google.com/sheets/api/samples/conditional-formatting
Conditional formatting: https://developers.google.com/sheets/api/guides/conditional-format
"""
function batch_update!(client::GoogleSheetsClient, spreadsheet::Spreadsheet, body::Dict)::Dict{Any,Any}
    @_print_python_exception begin
        sheet = client.client.spreadsheets()
        result = sheet.batchUpdate(spreadsheetId=spreadsheet.id, body=body).execute()
        return result
    end
end


"""
Adds a new sheet to a spreadsheet.
"""
function add_sheet!(client::GoogleSheetsClient, spreadsheet::Spreadsheet, title::AbstractString)::Dict{Any,Any}
    body = Dict(
        "requests" => [
            Dict(
                "addSheet" => Dict(
                    "properties" => Dict(
                        "title" => title,
                    )
                ),
            ),
        ],
    )

    return batch_update!(client, spreadsheet, body)
end


"""
Returns a dictionary of values to describe a 1D range of cells in a sheet.
"""
function cellRange1D(sheet_id::Int64, dim::AbstractString, start_index::Integer, end_index::Integer)
    return Dict(
        "sheetId" => sheet_id,
        "dimension" => dim,
        "startIndex" => start_index,
        "endIndex" => end_index,
    )
end


"""
Returns a dictionary of values to describe a 2D range of cells in a sheet.
"""
function cellRange2D(sheet_id::Int64, start_row_index::Integer, end_row_index::Integer, start_col_index::Integer, end_col_index::Integer)
    return Dict(
            "sheetId" => sheet_id,
            "startRowIndex" => start_row_index,
            "endRowIndex" => end_row_index+1,
            "startColumnIndex" => start_col_index,
            "endColumnIndex" => end_col_index+1,
        )
end


"""
Removes a sheet from a spreadsheet.
"""
function delete_sheet!(client::GoogleSheetsClient, spreadsheet::Spreadsheet, title::AbstractString)::Dict{Any,Any}
    properties = meta(client, spreadsheet, title)
    return delete_sheet!(client, spreadsheet, properties["sheetId"])
end


"""
Removes a sheet from a spreadsheet.
"""
function delete_sheet!(client::GoogleSheetsClient, spreadsheet::Spreadsheet, sheet_id::Int64)::Dict{Any,Any}
    body = Dict(
        "requests" => [
            Dict(
                "deleteSheet" => Dict(
                    "sheetId" => sheet_id,
                ),
            ),
        ],
    )

    return batch_update!(client, spreadsheet, body)
end


"""
Freeze rows and columns in a sheet.
"""
function freeze!(client::GoogleSheetsClient, spreadsheet::Spreadsheet, title::AbstractString, rows::Int64=0, cols::Int64=0)::Dict{Any,Any}
    properties = meta(client, spreadsheet, title)
    return freeze!(client, spreadsheet, properties["sheetId"], rows, cols)
end


"""
Freeze rows and columns in a sheet.
"""
function freeze!(client::GoogleSheetsClient, spreadsheet::Spreadsheet, sheet_id::Int64, rows::Int64=0, cols::Int64=0)::Dict{Any,Any}
    body = Dict(
        "requests" => [
            Dict(
                "updateSheetProperties" => Dict(
                    "properties" => Dict(
                        "sheetId" => sheet_id,
                        "gridProperties" => Dict(
                            "frozenRowCount" => rows,
                        ),
                    ),
                    "fields" => "gridProperties.frozenRowCount",
                ),
            ),

            Dict(
                "updateSheetProperties" => Dict(
                    "properties" => Dict(
                        "sheetId" => sheet_id,
                        "gridProperties" => Dict(
                            "frozenColumnCount" => cols,
                        ),
                    ),
                    "fields" => "gridProperties.frozenColumnCount",
                ),
            ),
        ],
    )

    return batch_update!(client, spreadsheet, body)
end


"""
Append rows and columns to a sheet.
"""
function Base.append!(client::GoogleSheetsClient, spreadsheet::Spreadsheet, title::AbstractString, rows::Int64=0, cols::Int64=0)::Dict{Any,Any}
    properties = meta(client, spreadsheet, title)
    return append!(client, spreadsheet, properties["sheetId"], rows, cols)
end


"""
Append rows and columns to a sheet.
"""
function Base.append!(client::GoogleSheetsClient, spreadsheet::Spreadsheet, sheet_id::Int64, rows::Int64=0, cols::Int64=0)::Dict{Any,Any}
    body = Dict(
        "requests" => [
            Dict(
                "appendDimension" => Dict(
                    "sheetId" => sheet_id,
                    "dimension" => "ROWS",
                    "length" => rows,
                ),
            ),

            Dict(
                "appendDimension" => Dict(
                    "sheetId" => sheet_id,
                    "dimension" => "COLUMNS",
                    "length" => cols,
                ),
            ),
        ],
    )

    return batch_update!(client, spreadsheet, body)
end


"""
Insert rows into to a sheet.
"""
function insert_rows!(client::GoogleSheetsClient, spreadsheet::Spreadsheet, title::AbstractString, start_index::Integer, end_index::Integer)::Dict{Any,Any}
    properties = meta(client, spreadsheet, title)
    return insert_rows!(client, spreadsheet, properties["sheetId"], start_index, end_index)
end


"""
Insert rows into to a sheet.
"""
function insert_rows!(client::GoogleSheetsClient, spreadsheet::Spreadsheet, sheet_id::Int64, start_index::Integer, end_index::Integer)::Dict{Any,Any}
    return _insert!(client, spreadsheet, sheet_id, "ROWS", start_index, end_index)
end


"""
Insert columns into to a sheet.
"""
function insert_cols!(client::GoogleSheetsClient, spreadsheet::Spreadsheet, title::AbstractString, start_index::Integer, end_index::Integer)::Dict{Any,Any}
    properties = meta(client, spreadsheet, title)
    return insert_cols!(client, spreadsheet, properties["sheetId"], start_index, end_index)
end


"""
Insert columns into to a sheet.
"""
function insert_cols!(client::GoogleSheetsClient, spreadsheet::Spreadsheet, sheet_id::Int64, start_index::Integer, end_index::Integer)::Dict{Any,Any}
    return _insert!(client, spreadsheet, sheet_id, "COLUMNS", start_index, end_index)
end


"""
Insert rows or columns into to a sheet.
"""
function _insert!(client::GoogleSheetsClient, spreadsheet::Spreadsheet, sheet_id::Int64, dim::AbstractString, start_index::Integer, end_index::Integer)::Dict{Any,Any}
    body = Dict(
        "requests" => [
            Dict(
                "insertDimension" => Dict(
                    "range" => cellRange1D(sheet_id, dim, start_index, end_index),
                ),
            ),
        ],
    )

    return batch_update!(client, spreadsheet, body)
end


"""
Delete rows from a sheet.
"""
function delete_rows!(client::GoogleSheetsClient, spreadsheet::Spreadsheet, title::AbstractString, start_index::Integer, end_index::Integer)::Dict{Any,Any}
    properties = meta(client, spreadsheet, title)
    return delete_rows!(client, spreadsheet, properties["sheetId"], start_index, end_index)
end


"""
Delete rows from a sheet.
"""
function delete_rows!(client::GoogleSheetsClient, spreadsheet::Spreadsheet, sheet_id::Int64, start_index::Integer, end_index::Integer)::Dict{Any,Any}
    return _delete!(client, spreadsheet, sheet_id, "ROWS", start_index, end_index)
end


"""
Delete columns from a sheet.
"""
function delete_cols!(client::GoogleSheetsClient, spreadsheet::Spreadsheet, title::AbstractString, start_index::Integer, end_index::Integer)::Dict{Any,Any}
    properties = meta(client, spreadsheet, title)
    return delete_cols!(client, spreadsheet, properties["sheetId"], start_index, end_index)
end


"""
Delete columns from a sheet.
"""
function delete_cols!(client::GoogleSheetsClient, spreadsheet::Spreadsheet, sheet_id::Int64, start_index::Integer, end_index::Integer)::Dict{Any,Any}
    return _delete!(client, spreadsheet, sheet_id, "COLUMNS", start_index, end_index)
end


"""
Delete rows or columns from a sheet.
"""
function _delete!(client::GoogleSheetsClient, spreadsheet::Spreadsheet, sheet_id::Int64, dim::AbstractString, start_index::Integer, end_index::Integer)::Dict{Any,Any}
    body = Dict(
        "requests" => [
            Dict(
                "deleteDimension" => Dict(
                    "range" => cellRange1D(sheet_id, dim, start_index, end_index),
                ),
            ),
        ],
    )

    return batch_update!(client, spreadsheet, body)
end


"""
Formats number values.

See: https://developers.google.com/sheets/api/guides/formats
"""
function format_number!(client::GoogleSheetsClient, spreadsheet::Spreadsheet, title::AbstractString, start_row_index::Integer, end_row_index::Integer, start_col_index::Integer, end_col_index::Integer, format_pattern::AbstractString)::Dict{Any,Any}
    properties = meta(client, spreadsheet, title)
    return format_number!(client, spreadsheet, properties["sheetId"], start_row_index, end_row_index, start_col_index, end_col_index, format_pattern)
end


"""
Formats number values.

See: https://developers.google.com/sheets/api/guides/formats
"""
function format_number!(client::GoogleSheetsClient, spreadsheet::Spreadsheet, sheet_id::Int64, start_row_index::Integer, end_row_index::Integer, start_col_index::Integer, end_col_index::Integer, format_pattern::AbstractString)::Dict{Any,Any}
    return _format_number!(client, spreadsheet, sheet_id, start_row_index, end_row_index, start_col_index, end_col_index, "NUMBER", format_pattern)
end


"""
Formats date-time values.

See: https://developers.google.com/sheets/api/guides/formats
"""
function format_datetime!(client::GoogleSheetsClient, spreadsheet::Spreadsheet, title::AbstractString, start_row_index::Integer, end_row_index::Integer, start_col_index::Integer, end_col_index::Integer, format_pattern::AbstractString)::Dict{Any,Any}
    properties = meta(client, spreadsheet, title)
    return format_datetime!(client, spreadsheet, properties["sheetId"], start_row_index, end_row_index, start_col_index, end_col_index, format_pattern)
end


"""
Formats date-time values.

See: https://developers.google.com/sheets/api/guides/formats
"""
function format_datetime!(client::GoogleSheetsClient, spreadsheet::Spreadsheet, sheet_id::Int64, start_row_index::Integer, end_row_index::Integer, start_col_index::Integer, end_col_index::Integer, format_pattern::AbstractString)::Dict{Any,Any}
    return _format_number!(client, spreadsheet, sheet_id, start_row_index, end_row_index, start_col_index, end_col_index, "DATE", format_pattern)
end


"""
Formats number values.

See: https://developers.google.com/sheets/api/guides/formats
"""
function _format_number!(client::GoogleSheetsClient, spreadsheet::Spreadsheet, sheet_id::Int64, start_row_index::Integer, end_row_index::Integer, start_col_index::Integer, end_col_index::Integer, format_type::AbstractString, format_pattern::AbstractString)::Dict{Any,Any}
    body = Dict(
        "requests" => [
            Dict(
                "repeatCell" => Dict(
                    "range" => cellRange2D(sheet_id, start_row_index, end_row_index, start_col_index, end_col_index),

                    "cell" => Dict(
                        "userEnteredFormat" => Dict(
                            "numberFormat" => Dict(
                                "type" => format_type,
                                "pattern" => format_pattern,
                            ),
                        ),
                    ),

                    "fields" => "userEnteredFormat.numberFormat",
                ),
            ),
        ],
    )

    return batch_update!(client, spreadsheet, body)
end


"""
Sets the background color.

See: https://developers.google.com/sheets/api/guides/formats
"""
function format_background_color!(client::GoogleSheetsClient, spreadsheet::Spreadsheet, title::AbstractString, start_row_index::Integer, end_row_index::Integer, start_col_index::Integer, end_col_index::Integer, color::Colorant)::Dict{Any,Any}
    properties = meta(client, spreadsheet, title)
    return format_background_color!(client, spreadsheet, properties["sheetId"], start_row_index, end_row_index, start_col_index, end_col_index, color)
end


"""
Sets the background color.

See: https://developers.google.com/sheets/api/guides/formats
"""
function format_background_color!(client::GoogleSheetsClient, spreadsheet::Spreadsheet, sheet_id::Int64, start_row_index::Integer, end_row_index::Integer, start_col_index::Integer, end_col_index::Integer, color::Colorant)::Dict{Any,Any}
    c = convert(RGBA, color)

    body = Dict(
        "requests" => [
            Dict(
                "repeatCell" => Dict(
                    "range" => cellRange2D(sheet_id, start_row_index, end_row_index, start_col_index, end_col_index),

                    "cell" => Dict(
                        "userEnteredFormat" => Dict(
                            "backgroundColor" => Dict(
                                "red" => red(c),
                                "green" => green(c),
                                "blue" => blue(c),
                                "alpha" => alpha(c),
                            ),
                        ),
                    ),

                    "fields" => "userEnteredFormat.backgroundColor",
                ),
            ),
        ],
    )

    return batch_update!(client, spreadsheet, body)
end

#TODO: add chart -> https://developers.google.com/sheets/api/reference/rest/v4/spreadsheets/charts
#TODO: add filterview -> batch_update -> https://developers.google.com/sheets/api/guides/filters
#TODO: add basic formatting -> https://developers.google.com/sheets/api/samples/formatting
#TODO: add conditional formatting -> https://developers.google.com/sheets/api/samples/conditional-formatting  https://developers.google.com/sheets/api/guides/conditional-format

end # module
