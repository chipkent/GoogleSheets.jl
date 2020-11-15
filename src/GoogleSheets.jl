
"""
A package for working with Google Sheets.
"""
module GoogleSheets

using PyCall

export AuthService, auth_service, Spreadsheet, CellRange, get, update

#TODO ??? to docs???
# main one -> https://developers.google.com/sheets/api/guides/concepts
# https://developers.google.com/sheets/api/quickstart/python
# https://developers.google.com/sheets/api/guides/authorizing


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
A permission for accessing Google resources.

AUTH_SPREADSHEET_READONLY:  Allows read-only access to the user's sheets and their properties.
AUTH_SPREADSHEET_READWRITE: Allows read/write access to the user's sheets and their properties.
"""
@exported_enum AuthPermission AUTH_SPREADSHEET_READONLY AUTH_SPREADSHEET_READWRITE


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
    id::String
end


"""
A range of cells within a spreadsheet.
"""
struct CellRange
    """Spreadsheet containing the cells."""
    spreadsheet::Spreadsheet

    """Range of cells."""
    range::String
end


"""
Maps permissions to the appropriate permission URLs.
"""
function scope_urls(permissions::AuthPermission)::Array{String,1}
    return [permission_urls[permissions]]
end


"""
Maps permissions to the appropriate permission URLs.
"""
function scope_urls(permissions::Array{AuthPermission,1})::Array{String,1}
    return [permission_urls[perm] for perm in permissions]
end


#TODO rename -> rename AUTH stuff to client???
#TODO handle token file locations
"""
Creates a client for accessing Google Sheets.
"""
function sheets_client(permissions::Union{AuthPermission,Array{AuthPermission,1}})::GoogleSheetsClient
    pickle = pyimport("pickle")
    os_path = pyimport("os.path")
    build = pyimport("googleapiclient.discovery").build
    InstalledAppFlow = pyimport("google_auth_oauthlib.flow").InstalledAppFlow
    Request = pyimport("google.auth.transport.requests").Request
    open = pybuiltin("open")

    scopes = scope_urls(permissions)
    creds = nothing

    # The file token.pickle stores the user's access and refresh tokens, and is
    # created automatically when the authorization flow completes for the first
    # time.
    if os_path.exists("token.pickle")
        @pywith open("token.pickle", "rb") as token begin
            creds = pickle.load(token)
        end
    end

    # If there are no (valid) credentials available, let the user log in.
    if isnothing(creds) || !creds.valid
        if !isnothing(creds) && creds.expired && !isnothing(creds.refresh_token)
            creds.refresh(Request())
        else
            flow = InstalledAppFlow.from_client_secrets_file("credentials.json", scopes)
            creds = flow.run_local_server(port=0)
        end

        # Save the credentials for the next run
        @pywith open("token.pickle", "wb") as token begin
            pickle.dump(creds, token)
        end
    end

    return GoogleSheetsClient(build("sheets", "v4", credentials=creds))
end


"""
Gets a range of cell values from a spreadsheet.
"""
function Base.get(client::GoogleSheetsClient, range::CellRange)::Dict{Any,Any}
    sheet = client.client.spreadsheets()
    result = sheet.values().get(spreadsheetId=range.spreadsheet.id,
                                majorDimension="ROWS",
                                range=range.range).execute()

    #TODO return a struct?? with a cell range???
    return result
end

#TODO add put! update?
#TODO return value type???
#TODO document raw
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
function update(client::GoogleSheetsClient, range::CellRange, values::Array{<:Any,2}; raw::Bool=false)::Dict{Any,Any}
    body = Dict(
        "values" => values,
        "majorDimension" => "ROWS",
    )

    sheet = client.client.spreadsheets()
    result = sheet.values().update(spreadsheetId=range.spreadsheet.id,
                                range=range.range,
                                valueInputOption= raw ? "RAW" : "USER_ENTERED",
                                body=body).execute()

    #TODO return a struct?? with a cell range???
    return result
end

end # module
