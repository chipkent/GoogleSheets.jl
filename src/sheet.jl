export freeze!, append!, insert_rows!, insert_cols!, delete_rows!, delete_cols!


"""
Freeze rows and columns in a sheet.
"""
function freeze!(client::GoogleSheetsClient, sheet::Sheet, rows::Int64=0, cols::Int64=0)::Dict{Any,Any}
    body = Dict(
        "requests" => [
            Dict(
                "updateSheetProperties" => Dict(
                    "properties" => Dict(
                        "sheetId" => sheet.id,
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
                        "sheetId" => sheet.id,
                        "gridProperties" => Dict(
                            "frozenColumnCount" => cols,
                        ),
                    ),
                    "fields" => "gridProperties.frozenColumnCount",
                ),
            ),
        ],
    )

    return batch_update!(client, sheet.spreadsheet, body)
end


"""
Append rows and columns to a sheet.
"""
function Base.append!(client::GoogleSheetsClient, sheet::Sheet, rows::Int64=0, cols::Int64=0)::Dict{Any,Any}
    body = Dict(
        "requests" => [
            Dict(
                "appendDimension" => Dict(
                    "sheetId" => sheet.id,
                    "dimension" => "ROWS",
                    "length" => rows,
                ),
            ),

            Dict(
                "appendDimension" => Dict(
                    "sheetId" => sheet.id,
                    "dimension" => "COLUMNS",
                    "length" => cols,
                ),
            ),
        ],
    )

    return batch_update!(client, sheet.spreadsheet, body)
end


"""
Insert rows into to a sheet.
"""
function insert_rows!(client::GoogleSheetsClient, sheet::Sheet, start_index::Integer, end_index::Integer)::Dict{Any,Any}
    return _insert!(client, sheet, "ROWS", start_index, end_index)
end


"""
Insert columns into to a sheet.
"""
function insert_cols!(client::GoogleSheetsClient, sheet::Sheet, start_index::Integer, end_index::Integer)::Dict{Any,Any}
    return _insert!(client, sheet, "COLUMNS", start_index, end_index)
end


"""
Insert rows or columns into to a sheet.
"""
function _insert!(client::GoogleSheetsClient, sheet::Sheet, dim::AbstractString, start_index::Integer, end_index::Integer)::Dict{Any,Any}
    body = Dict(
        "requests" => [
            Dict(
                "insertDimension" => Dict(
                    "range" => cellRange1D(sheet, dim, start_index, end_index),
                ),
            ),
        ],
    )

    return batch_update!(client, sheet.spreadsheet, body)
end


"""
Delete rows from a sheet.
"""
function delete_rows!(client::GoogleSheetsClient, sheet::Sheet, start_index::Integer, end_index::Integer)::Dict{Any,Any}
    return _delete!(client, sheet, "ROWS", start_index, end_index)
end


"""
Delete columns from a sheet.
"""
function delete_cols!(client::GoogleSheetsClient, sheet::Sheet, start_index::Integer, end_index::Integer)::Dict{Any,Any}
    return _delete!(client, sheet, "COLUMNS", start_index, end_index)
end


"""
Delete rows or columns from a sheet.
"""
function _delete!(client::GoogleSheetsClient, sheet::Sheet, dim::AbstractString, start_index::Integer, end_index::Integer)::Dict{Any,Any}
    body = Dict(
        "requests" => [
            Dict(
                "deleteDimension" => Dict(
                    "range" => cellRange1D(sheet, dim, start_index, end_index),
                ),
            ),
        ],
    )

    return batch_update!(client, sheet.spreadsheet, body)
end