
using PyCall

"""
A package for working with Google Sheets.
"""
module GoogleSheets

export AuthService, auth_service, Spreadsheet, Range, get

#TODO ??? to docs???
# https://developers.google.com/sheets/api/quickstart/python

struct AuthService
    service
end

struct Spreadsheet
    id::String
end

struct Range
    spreadsheet::Spreadsheet
    range::String
end

#TODO rename
#TODO document
function auth_service()::AuthService
    pickle = pyimport("pickle")
    os_path = pyimport("os.path")
    build = pyimport("googleapiclient.discovery").build
    InstalledAppFlow = pyimport("google_auth_oauthlib.flow").InstalledAppFlow
    Request = pyimport("google.auth.transport.requests").Request
    open = pybuiltin("open")

    #TODO scopes selection
    #TODO map scope to tokens
    #TODO @enum to make an enum ---> right or wrong way??? or a class with stuff in it? or sym?
    # If modifying these scopes, delete the file token.pickle.
    SCOPES = ["https://www.googleapis.com/auth/spreadsheets.readonly"]

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
            flow = InstalledAppFlow.from_client_secrets_file("credentials.json", SCOPES)
            creds = flow.run_local_server(port=0)
        end

        # Save the credentials for the next run
        @pywith open("token.pickle", "wb") as token begin
            pickle.dump(creds, token)
        end
    end

    return AuthService(build("sheets", "v4", credentials=creds))
end

"""
Gets a range of values from a spreadsheet.
"""
function Base.get(auth::AuthService, range::Range)::Dict{Any,Any}
    sheet = auth.service.spreadsheets()
    result = sheet.values().get(spreadsheetId=range.spreadsheet.id,
                                range=range.range).execute()

    #TODO return a struct??
    return result
end

#TODO delete me!!!
# service = auth_service()
#
# # The ID and range of a sample spreadsheet.
# SAMPLE_SPREADSHEET_ID = "1BxiMVs0XRA5nFMdKvBdBZjgmUUqptlbs74OgvE2upms"
# SAMPLE_RANGE_NAME = "Class Data!A2:E"
#
# # Call the Sheets API
# sheet = service.spreadsheets()
# result = sheet.values().get(spreadsheetId=SAMPLE_SPREADSHEET_ID,
#                             range=SAMPLE_RANGE_NAME).execute()
#
# println("KEYS: $(keys(result))")
# println("RANGE: $(result["range"])")
# println("MAJORDIM: $(result["majorDimension"])")
#
# values = result["values"]
#
# if isnothing(values)
#     println("No data found.")
# else
#     println("Name, Major:")
#     for row in eachrow(values)
#         # Print columns A and E, which correspond to indices 0 and 4.
#         println("ROW: $row")
#         # println("%s, %s" % (row[0], row[4]))
#     end
# end
#
# greet() = print("Hello World!")

end # module
