export meta, show


"""
Gets metadata about a spreadsheet.
"""
function meta(client::GoogleSheetsClient, spreadsheet::Spreadsheet)::Dict{Any,Any}
    return gsapi_spreadsheet_get(client; spreadsheetId=spreadsheet.id)
end


"""
Gets metadata about a spreadsheet key-value pair.
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
Gets metadata about a spreadsheet sheet.
"""
meta(client::GoogleSheetsClient, spreadsheet::Spreadsheet, title::AbstractString)::Dict{Any,Any} = meta(client, spreadsheet, "title", title)


"""
Gets metadata about a spreadsheet sheet.
"""
meta(client::GoogleSheetsClient, spreadsheet::Spreadsheet, sheet_id::Int64)::Dict{Any,Any} = meta(client, spreadsheet, "sheetId", sheet_id)


"""
Gets metadata about a spreadsheet sheet.
"""
meta(client::GoogleSheetsClient, sheet::Sheet)::Dict{Any,Any} = meta(client, sheet.spreadsheet, sheet.id)


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
Prints metadata about a spreadsheet sheet.
"""
Base.show(client::GoogleSheetsClient, sheet::Sheet) = Base.show(client, sheet.spreadsheet, sheet.id)


