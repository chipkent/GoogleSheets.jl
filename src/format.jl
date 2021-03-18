export format_number!, format_datetime!, format_background_color!


"""
Formats number values.

See: https://developers.google.com/sheets/api/guides/formats
"""
function format_number!(client::GoogleSheetsClient, spreadsheet::Spreadsheet, title::AbstractString, start_row_index::Integer, end_row_index::Integer, start_col_index::Integer, end_col_index::Integer, format_pattern::AbstractString)::Dict{Any,Any}
     #TODO: get rid of this stuff???
     properties = meta(client, spreadsheet, title)
    return format_number!(client, spreadsheet, properties["sheetId"], start_row_index, end_row_index, start_col_index, end_col_index, format_pattern)
end


"""
Formats number values.

See: https://developers.google.com/sheets/api/guides/formats
"""
function format_number!(client::GoogleSheetsClient, spreadsheet::Spreadsheet, sheet_id::Int64, start_row_index::Integer, end_row_index::Integer, start_col_index::Integer, end_col_index::Integer, format_pattern::AbstractString)::Dict{Any,Any}
    return _format_number!(client, spreadsheet, sheet_id, start_row_index, end_row_index, start_col_index, end_col_index, "NUMBER", format_pattern)
end


"""
Formats date-time values.

See: https://developers.google.com/sheets/api/guides/formats
"""
function format_datetime!(client::GoogleSheetsClient, spreadsheet::Spreadsheet, title::AbstractString, start_row_index::Integer, end_row_index::Integer, start_col_index::Integer, end_col_index::Integer, format_pattern::AbstractString)::Dict{Any,Any}
     #TODO: get rid of this stuff???
     properties = meta(client, spreadsheet, title)
    return format_datetime!(client, spreadsheet, properties["sheetId"], start_row_index, end_row_index, start_col_index, end_col_index, format_pattern)
end


"""
Formats date-time values.

See: https://developers.google.com/sheets/api/guides/formats
"""
function format_datetime!(client::GoogleSheetsClient, spreadsheet::Spreadsheet, sheet_id::Int64, start_row_index::Integer, end_row_index::Integer, start_col_index::Integer, end_col_index::Integer, format_pattern::AbstractString)::Dict{Any,Any}
    return _format_number!(client, spreadsheet, sheet_id, start_row_index, end_row_index, start_col_index, end_col_index, "DATE", format_pattern)
end


"""
Formats number values.

See: https://developers.google.com/sheets/api/guides/formats
"""
function _format_number!(client::GoogleSheetsClient, spreadsheet::Spreadsheet, sheet_id::Int64, start_row_index::Integer, end_row_index::Integer, start_col_index::Integer, end_col_index::Integer, format_type::AbstractString, format_pattern::AbstractString)::Dict{Any,Any}
    body = Dict(
        "requests" => [
            Dict(
                "repeatCell" => Dict(
                    "range" => cellRange2D(sheet_id, start_row_index, end_row_index, start_col_index, end_col_index),

                    "cell" => Dict(
                        "userEnteredFormat" => Dict(
                            "numberFormat" => Dict(
                                "type" => format_type,
                                "pattern" => format_pattern,
                            ),
                        ),
                    ),

                    "fields" => "userEnteredFormat.numberFormat",
                ),
            ),
        ],
    )

    return batch_update!(client, spreadsheet, body)
end


"""
Sets the background color.

See: https://developers.google.com/sheets/api/guides/formats
"""
function format_background_color!(client::GoogleSheetsClient, spreadsheet::Spreadsheet, title::AbstractString, start_row_index::Integer, end_row_index::Integer, start_col_index::Integer, end_col_index::Integer, color::Colorant)::Dict{Any,Any}
    #TODO: get rid of this stuff???
    properties = meta(client, spreadsheet, title)
    return format_background_color!(client, spreadsheet, properties["sheetId"], start_row_index, end_row_index, start_col_index, end_col_index, color)
end


"""
Sets the background color.

See: https://developers.google.com/sheets/api/guides/formats
"""
function format_background_color!(client::GoogleSheetsClient, spreadsheet::Spreadsheet, sheet_id::Int64, start_row_index::Integer, end_row_index::Integer, start_col_index::Integer, end_col_index::Integer, color::Colorant)::Dict{Any,Any}
    body = Dict(
        "requests" => [
            Dict(
                "repeatCell" => Dict(
                    "range" => cellRange2D(sheet_id, start_row_index, end_row_index, start_col_index, end_col_index),

                    "cell" => Dict(
                        "userEnteredFormat" => Dict(
                            "backgroundColor" => gsheet_json(color),
                        ),
                    ),

                    "fields" => "userEnteredFormat.backgroundColor",
                ),
            ),
        ],
    )

    return batch_update!(client, spreadsheet, body)
end