
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
import MacroTools

export GoogleSheetsClient, Spreadsheet, CellRange, sheets_client, meta, get, update!,
        batch_update!, add_sheet!, delete_sheet!


"""
Directory containing configuration files.
"""
config_dir = joinpath(homedir(),".julia/config/google_sheets/")


"""
Create an enum and export the enum type and all values.
"""
macro exported_enum(name, args...)
    return esc(quote
        @enum($name, $(args...))
        export $name
        $([:(export $arg) for arg in args]...)
    end)
end

"""
Print details on a python exception.
"""
macro print_python_exception(ex)
    # MacroTools.@q is used instead of quote so that the returned stacktrace
    # has line numbers from the calling function and not the macro.
    return esc(MacroTools.@q begin
        try
            $ex
        catch e
            #TODO how to fix this???
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
@exported_enum AuthScope AUTH_SPREADSHEET_READONLY AUTH_SPREADSHEET_READWRITE


const permission_urls = Dict(
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
Maps authorization scopes to the appropriate permission URLs.
"""
function scope_urls(scopes::AuthScope)::Array{String,1}
    return [permission_urls[scopes]]
end


"""
Maps authorization scopes to the appropriate permission URLs.
"""
function scope_urls(scopes::Array{AuthScope,1})::Array{String,1}
    return [permission_urls[scope] for scope in scopes]
end


"""
Gets the credentials file needed to log into Google.
See the python quick start reference for a link to generate credentials.
https://developers.google.com/sheets/api/quickstart/python
"""
function credentials_file()::String
    file = "credentials.json"
    return joinpath(config_dir, file)
end


"""
Gets a cached token file which allows access to Google Sheets with
specific authorization scopes.
"""
function token_file(scopes::AuthScope)::String
    return token_file([scopes])
end


"""
Gets a cached token file which allows access to Google Sheets with
specific authorization scopes.
"""
function token_file(scopes::Array{AuthScope,1})::String
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
    tokenFile = token_file(scopes)
    scopeUrls = scope_urls(scopes)

    @print_python_exception begin
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
    @print_python_exception begin
        sheet = client.client.spreadsheets()
        result = sheet.get(spreadsheetId=spreadsheet.id).execute()

        #TODO return a struct?? with a cell range???
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
Gets a range of cell values from a spreadsheet.
"""
function Base.get(client::GoogleSheetsClient, range::CellRange)::Dict{Any,Any}
    @print_python_exception begin
        sheet = client.client.spreadsheets()
        result = sheet.values().get(spreadsheetId=range.spreadsheet.id,
                                    majorDimension="ROWS",
                                    range=range.range).execute()

        #TODO return a struct?? with a cell range???
        return result
    end
end


#TODO add put! update?
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
function update!(client::GoogleSheetsClient, range::CellRange, values::Array{<:Any,2}; raw::Bool=false)::Dict{Any,Any}
    body = Dict(
        "values" => values,
        "majorDimension" => "ROWS",
    )

    @print_python_exception begin
        sheet = client.client.spreadsheets()
        result = sheet.values().update(spreadsheetId=range.spreadsheet.id,
                                range=range.range,
                                valueInputOption= raw ? "RAW" : "USER_ENTERED",
                                body=body).execute()

        #TODO return a struct?? with a cell range???
        return result
    end
end


"""
Applies one or more updates to a spreadsheet.

Each request is validated before being applied. If any request is not valid then
the entire request will fail and nothing will be applied.
"""
function batch_update!(client::GoogleSheetsClient, spreadsheet::Spreadsheet, body::Dict)::Dict{Any,Any}
    @print_python_exception begin
        sheet = client.client.spreadsheets()
        result = sheet.batchUpdate(spreadsheetId=spreadsheet.id, body=body).execute()

        #TODO return a struct?? with a cell range???
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
Removes a sheet from a spreadsheet.
"""
function delete_sheet!(client::GoogleSheetsClient, spreadsheet::Spreadsheet, title::AbstractString)::Dict{Any,Any}
    properties = meta(client, spreadsheet, title)
    return delete_sheet!(client, spreadsheet, properties["sheetId"])
end


"""
Removes a sheet from a spreadsheet.
"""
function delete_sheet!(client::GoogleSheetsClient, spreadsheet::Spreadsheet, sheetId::Int64)::Dict{Any,Any}
    body = Dict(
        "requests" => [
            Dict(
                "deleteSheet" => Dict(
                    "sheetId" => sheetId,
                ),
            ),
        ],
    )

    return batch_update!(client, spreadsheet, body)
end


#TODO batchGet
#TODO append
#TODO add chart
#TODO add filterview
#TODO add conditional formatting
#TODO insert rows
#TODO delete rows

end # module
