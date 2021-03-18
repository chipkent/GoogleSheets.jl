export freeze!, append!, insert_rows!, insert_cols!, delete_rows!, delete_cols!


"""
Freeze rows and columns in a sheet.
"""
function freeze!(client::GoogleSheetsClient, spreadsheet::Spreadsheet, title::AbstractString, rows::Int64=0, cols::Int64=0)::Dict{Any,Any}
     #TODO: get rid of this stuff???
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
    #TODO: get rid of this stuff???
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
     #TODO: get rid of this stuff???
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
     #TODO: get rid of this stuff???
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
     #TODO: get rid of this stuff???
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
    #TODO: get rid of this stuff???
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