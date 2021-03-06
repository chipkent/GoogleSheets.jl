export meta, show


"""
    meta(client::GoogleSheetsClient, spreadsheet::Spreadsheet)::Dict{Any,Any}

Gets metadata about a spreadsheet.

# Arguments
- `client::GoogleSheetsClient`: client
- `spreadsheet::Spreadsheet`: spreadsheet
"""
function meta(client::GoogleSheetsClient, spreadsheet::Spreadsheet)::Dict{Any,Any}
    return gsheet_api_spreadsheet_get(client; spreadsheetId=spreadsheet.id)
end


"""
    meta(client::GoogleSheetsClient, spreadsheet::Spreadsheet, key::AbstractString, value)::Dict{Any,Any}

Gets metadata about a spreadsheet key-value pair.

# Arguments
- `client::GoogleSheetsClient`: client
- `spreadsheet::Spreadsheet`: spreadsheet
- `key:AbstractString`: key
- `value`: value
"""
function meta(client::GoogleSheetsClient, spreadsheet::Spreadsheet, key::AbstractString, value)::Dict{Any,Any}
    metadata = meta(client, spreadsheet)
    sheets = metadata["sheets"]

    for sheet in sheets
        properties = sheet["properties"]

        if properties[key] == value
            return properties
        end
    end

    throw(KeyError(title))
end


"""
    meta(client::GoogleSheetsClient, spreadsheet::Spreadsheet, title::AbstractString)::Dict{Any,Any}

Gets metadata about a spreadsheet sheet.

# Arguments
- `client::GoogleSheetsClient`: client
- `spreadsheet::Spreadsheet`: spreadsheet
- `title::AbstractString`: sheet title
"""
meta(client::GoogleSheetsClient, spreadsheet::Spreadsheet, title::AbstractString)::Dict{Any,Any} = meta(client, spreadsheet, "title", title)


"""
    meta(client::GoogleSheetsClient, spreadsheet::Spreadsheet, sheet_id::Int64)::Dict{Any,Any}

Gets metadata about a spreadsheet sheet.

# Arguments
- `client::GoogleSheetsClient`: client
- `spreadsheet::Spreadsheet`: spreadsheet
- `sheet_id::Int64`: sheet id
"""
meta(client::GoogleSheetsClient, spreadsheet::Spreadsheet, sheet_id::Int64)::Dict{Any,Any} = meta(client, spreadsheet, "sheetId", sheet_id)


"""
    meta(client::GoogleSheetsClient, sheet::Sheet)::Dict{Any,Any}

Gets metadata about a spreadsheet sheet.

# Arguments
- `client::GoogleSheetsClient`: client
- `sheet::Sheet`: sheet
"""
meta(client::GoogleSheetsClient, sheet::Sheet)::Dict{Any,Any} = meta(client, sheet.spreadsheet, sheet.id)


"""
    show(client::GoogleSheetsClient, spreadsheet::Spreadsheet)

Prints metadata about a spreadsheet.

# Arguments
- `client::GoogleSheetsClient`: client
- `spreadsheet::Spreadsheet`: spreadsheet
"""
function Base.show(client::GoogleSheetsClient, spreadsheet::Spreadsheet)
    m = meta(client, spreadsheet)
    println("Spreadsheet:")
    println(json(m, 4))
end


"""
    show(client::GoogleSheetsClient, spreadsheet::Spreadsheet, title::AbstractString)

Prints metadata about a spreadsheet sheet.

# Arguments
- `client::GoogleSheetsClient`: client
- `spreadsheet::Spreadsheet`: spreadsheet
- `title::AbstractString`: sheet title
"""
function Base.show(client::GoogleSheetsClient, spreadsheet::Spreadsheet, title::AbstractString)
    m = meta(client, spreadsheet, title)
    println("Sheet:")
    println(json(m, 4))
end


"""
    show(client::GoogleSheetsClient, spreadsheet::Spreadsheet, sheet_id::Int64)

Prints metadata about a spreadsheet sheet.

# Arguments
- `client::GoogleSheetsClient`: client
- `spreadsheet::Spreadsheet`: spreadsheet
- `sheet_id::Int64`: sheet id
"""
function Base.show(client::GoogleSheetsClient, spreadsheet::Spreadsheet, sheet_id::Int64)
    m = meta(client, spreadsheet, sheet_id)
    println("Sheet:")
    println(json(m, 4))
end


"""
    show(client::GoogleSheetsClient, sheet::Sheet)

Prints metadata about a spreadsheet sheet.

# Arguments
- `client::GoogleSheetsClient`: client
- `sheet::Sheet`: sheet
"""
Base.show(client::GoogleSheetsClient, sheet::Sheet) = Base.show(client, sheet.spreadsheet, sheet.id)


