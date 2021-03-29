export meta, show, sheet_names, sheets


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


#TODO: move to sheet?
"""
Gets the names of the sheets in the spreadsheet.
"""
function sheet_names(client::GoogleSheetsClient, spreadsheet::Spreadsheet)::Vector{String}
    m = meta(client, spreadsheet)
    return [ s["properties"]["title"] for s in m["sheets"] ]
end


#TODO: move to sheet?
"""
Gets the sheets in the spreadsheet.
"""
function sheets(client::GoogleSheetsClient, spreadsheet::Spreadsheet)::Vector{Sheet}
    m = meta(client, spreadsheet)
    return [ Sheet(spreadsheet, s["properties"]["sheetId"], s["properties"]["title"]) for s in m["sheets"] ]
end

