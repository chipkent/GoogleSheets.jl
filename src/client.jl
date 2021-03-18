export sheets_client


"""
Directory containing configuration files.
"""
config_dir = joinpath(homedir(),".julia/config/google_sheets/")


const _permission_urls = Dict(
    AUTH_SCOPE_READONLY => "https://www.googleapis.com/auth/spreadsheets.readonly",
    AUTH_SCOPE_READWRITE => "https://www.googleapis.com/auth/spreadsheets",
)


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