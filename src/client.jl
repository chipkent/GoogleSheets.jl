export sheets_client


# The Google Sheets API has access limits.
# The default limits are 100 requests per 100 seconds per user.
# Limits for reads and writes are tracked separately.
# Higher limits are available in the Google Cloud console.
# https://developers.google.com/sheets/api/limits
#
# Reality seems to be more complex.
# The tokens per second seems to match the described limit.
# If the max tokens are too large, the burstiness triggers limits, 
# without clearly indicating that burstiness is the problem.
#
# These defaults seem to work.
default_rate_limiter_tokens_per_sec = 0.95
default_rate_limiter_max_tokens = 5
default_rate_limiter_read = TokenBucketRateLimiter(default_rate_limiter_tokens_per_sec, default_rate_limiter_max_tokens, default_rate_limiter_max_tokens)
default_rate_limiter_write = TokenBucketRateLimiter(default_rate_limiter_tokens_per_sec, default_rate_limiter_max_tokens, default_rate_limiter_max_tokens)


"""
Update the default rate limiter.
"""
function update_default_rate_limiter(rate_limiter_tokens_per_sec::Float64; rate_limiter_max_tokens::Float64=5)
    global default_rate_limiter_read = TokenBucketRateLimiter(rate_limiter_tokens_per_sec, rate_limiter_max_tokens, rate_limiter_max_tokens)
    global default_rate_limiter_write = TokenBucketRateLimiter(rate_limiter_tokens_per_sec, rate_limiter_max_tokens, rate_limiter_max_tokens)    
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
function sheets_client(scopes::Union{AuthScope,Array{AuthScope,1}}; 
        rate_limiter_read::AbstractRateLimiter=default_rate_limiter_read, 
        rate_limiter_write::AbstractRateLimiter=default_rate_limiter_write)::GoogleSheetsClient

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

        return GoogleSheetsClient(build("sheets", "v4", credentials=creds), rate_limiter_read, rate_limiter_write)
    end
end


#TODO: rename all of these to _gsapi_<func>
"""
Calls the python get API.
"""
function _get(client::GoogleSheetsClient; kwargs...)
    @rate_limit client.rate_limiter_read 1 @_print_python_exception begin
        return client.client.spreadsheets().values().get(;kwargs...).execute()
    end
end


#TODO: is the values call even needed???
"""
Calls the python get API.
"""
function _get_novalues(client::GoogleSheetsClient; kwargs...)
    @rate_limit client.rate_limiter_read 1 @_print_python_exception begin
        return client.client.spreadsheets().get(;kwargs...).execute()
    end
end


"""
Calls the python batchGet API.
"""
function _batchGet(client::GoogleSheetsClient; kwargs...)
    @rate_limit client.rate_limiter_read 1 @_print_python_exception begin
        return client.client.spreadsheets().values().batchGet(;kwargs...).execute()
    end
end


"""
Calls the python update API.
"""
function _update(client::GoogleSheetsClient; kwargs...)
    @rate_limit client.rate_limiter_write 1 @_print_python_exception begin
        return client.client.spreadsheets().values().update(;kwargs...).execute()
    end
end


#TODO: is the values call even needed???
"""
Calls the python batchUpdate API.
"""
function _batchUpdate_novalues(client::GoogleSheetsClient; kwargs...)
    @rate_limit client.rate_limiter_write 1 @_print_python_exception begin
        return client.client.spreadsheets().batchUpdate(;kwargs...).execute()
    end
end
