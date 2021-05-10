export freeze!, append!, insert_rows!, insert_cols!, delete_rows!, delete_cols!


"""
    freeze!(client::GoogleSheetsClient, sheet::Sheet, rows::Int64=0, cols::Int64=0)::Dict{Any,Any}

Freeze rows and columns in a sheet.

# Arguments
- `client::GoogleSheetsClient`: client
- `sheet::Sheet`: sheet
- `rows::Int64=0`: number of rows
- `cols::Int64=0`: number of columns
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
    append!(client::GoogleSheetsClient, sheet::Sheet, rows::Int64=0, cols::Int64=0)::Dict{Any,Any}

Append rows and columns to a sheet.

# Arguments
- `client::GoogleSheetsClient`: client
- `sheet::Sheet`: sheet
- `rows::Int64=0`: number of rows
- `cols::Int64=0`: number of columns
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
    insert_rows!(client::GoogleSheetsClient, range::CellIndexRange1D)::Dict{Any,Any}

Insert rows into to a sheet.

# Arguments
- `client::GoogleSheetsClient`: client
- `range::CellIndexRange1D`: cell index range
"""
function insert_rows!(client::GoogleSheetsClient, range::CellIndexRange1D)::Dict{Any,Any}
    return _insert!(client, range, "ROWS")
end


"""
    insert_cols!(client::GoogleSheetsClient, range::CellIndexRange1D)::Dict{Any,Any}

Insert columns into to a sheet.

# Arguments
- `client::GoogleSheetsClient`: client
- `range::CellIndexRange1D`: cell index range
"""
function insert_cols!(client::GoogleSheetsClient, range::CellIndexRange1D)::Dict{Any,Any}
    return _insert!(client, range, "COLUMNS")
end


"""
Insert rows or columns into to a sheet.

# Arguments
- `client::GoogleSheetsClient`: client
- `range::CellIndexRange1D`: cell index range
- `dim::AbstractString`: dimension string
"""
function _insert!(client::GoogleSheetsClient, range::CellIndexRange1D, dim::AbstractString)::Dict{Any,Any}
    body = Dict(
        "requests" => [
            Dict(
                "insertDimension" => Dict(
                    "range" => gsheet_json(range, dim)
                ),
            ),
        ],
    )

    return batch_update!(client, range.sheet.spreadsheet, body)
end


"""
Delete rows from a sheet.

# Arguments
- `client::GoogleSheetsClient`: client
- `range::CellIndexRange1D`: cell index range
"""
function delete_rows!(client::GoogleSheetsClient, range::CellIndexRange1D)::Dict{Any,Any}
    return _delete!(client, range, "ROWS")
end


"""
    delete_cols!(client::GoogleSheetsClient, range::CellIndexRange1D)::Dict{Any,Any}

Delete columns from a sheet.

# Arguments
- `client::GoogleSheetsClient`: client
- `range::CellIndexRange1D`: cell index range
"""
function delete_cols!(client::GoogleSheetsClient, range::CellIndexRange1D)::Dict{Any,Any}
    return _delete!(client, range, "COLUMNS")
end


"""
Delete rows or columns from a sheet.

# Arguments
- `client::GoogleSheetsClient`: client
- `range::CellIndexRange1D`: cell index range
- `dim::AbstractString`: dimension string
"""
function _delete!(client::GoogleSheetsClient, range::CellIndexRange1D, dim::AbstractString)::Dict{Any,Any}
    body = Dict(
        "requests" => [
            Dict(
                "deleteDimension" => Dict(
                    "range" => gsheet_json(range, dim)
                ),
            ),
        ],
    )

    return batch_update!(client, range.sheet.spreadsheet, body)
end