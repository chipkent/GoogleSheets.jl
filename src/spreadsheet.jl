export sheet_names, sheets, batch_update!, add_sheet!, delete_sheet!


"""
    sheet_names(client::GoogleSheetsClient, spreadsheet::Spreadsheet)::Vector{String}

Gets the names of the sheets in the spreadsheet.

# Arguments
- `client::GoogleSheetsClient`: client
- `spreadsheet::Spreadsheet`: spreadsheet
"""
function sheet_names(client::GoogleSheetsClient, spreadsheet::Spreadsheet)::Vector{String}
    m = meta(client, spreadsheet)
    return [ s["properties"]["title"] for s in m["sheets"] ]
end


"""
    sheets(client::GoogleSheetsClient, spreadsheet::Spreadsheet)::Vector{Sheet}

Gets the sheets in the spreadsheet.

# Arguments
- `client::GoogleSheetsClient`: client
- `spreadsheet::Spreadsheet`: spreadsheet
"""
function sheets(client::GoogleSheetsClient, spreadsheet::Spreadsheet)::Vector{Sheet}
    m = meta(client, spreadsheet)
    return [ Sheet(spreadsheet, s["properties"]["sheetId"], s["properties"]["title"]) for s in m["sheets"] ]
end


"""
    batch_update!(client::GoogleSheetsClient, spreadsheet::Spreadsheet, body::Dict)::Dict{Any,Any}

Applies one or more updates to a spreadsheet.

Each request is validated before being applied. If any request is not valid then
the entire request will fail and nothing will be applied.

Common batch_update! functionality:
Charts: https://developers.google.com/sheets/api/reference/rest/v4/spreadsheets/charts
Filters: https://developers.google.com/sheets/api/guides/filters
Basic formatting: https://developers.google.com/sheets/api/samples/formatting
Conditional formatting: https://developers.google.com/sheets/api/samples/conditional-formatting
Conditional formatting: https://developers.google.com/sheets/api/guides/conditional-format

# Arguments
- `client::GoogleSheetsClient`: client
- `spreadsheet::Spreadsheet`: spreadsheet
- `body::Dict`: updte body
"""
function batch_update!(client::GoogleSheetsClient, spreadsheet::Spreadsheet, body::Dict)::Dict{Any,Any}
    return gsheet_api_speadsheet_batchupdate(client; spreadsheetId=spreadsheet.id, body=body)
end


"""
    add_sheet!(client::GoogleSheetsClient, spreadsheet::Spreadsheet, title::AbstractString)::Dict{Any,Any}

Adds a new sheet to a spreadsheet.

# Arguments
- `client::GoogleSheetsClient`: client
- `spreadsheet::Spreadsheet`: spreadsheet
- `title::AbstractString`: sheet title
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


#TODO: make this use sheet??
"""
    delete_sheet!(client::GoogleSheetsClient, spreadsheet::Spreadsheet, title::AbstractString)::Dict{Any,Any}

Removes a sheet from a spreadsheet.

# Arguments
- `client::GoogleSheetsClient`: client
- `spreadsheet::Spreadsheet`: spreadsheet
- `title::AbstractString`: sheet title
"""
function delete_sheet!(client::GoogleSheetsClient, spreadsheet::Spreadsheet, title::AbstractString)::Dict{Any,Any}
    #TODO: get rid of this stuff???
    properties = meta(client, spreadsheet, title)
    return delete_sheet!(client, spreadsheet, properties["sheetId"])
end


#TODO: make this use sheet??
"""
    delete_sheet!(client::GoogleSheetsClient, spreadsheet::Spreadsheet, sheet_id::Int64)::Dict{Any,Any}

Removes a sheet from a spreadsheet.

# Arguments
- `client::GoogleSheetsClient`: client
- `spreadsheet::Spreadsheet`: spreadsheet
- `sheet_id::Int64`: sheet id
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