export format_number!, format_datetime!, format_background_color!


"""
Formats number values.

See: https://developers.google.com/sheets/api/guides/formats
"""
function format_number!(client::GoogleSheetsClient, sheet::Sheet, start_row_index::Integer, end_row_index::Integer, start_col_index::Integer, end_col_index::Integer, format_pattern::AbstractString)::Dict{Any,Any}
    return _format_number!(client, sheet, start_row_index, end_row_index, start_col_index, end_col_index, "NUMBER", format_pattern)
end


"""
Formats date-time values.

See: https://developers.google.com/sheets/api/guides/formats
"""
function format_datetime!(client::GoogleSheetsClient, sheet::Sheet, start_row_index::Integer, end_row_index::Integer, start_col_index::Integer, end_col_index::Integer, format_pattern::AbstractString)::Dict{Any,Any}
    return _format_number!(client, sheet, start_row_index, end_row_index, start_col_index, end_col_index, "DATE", format_pattern)
end


"""
Formats number values.

See: https://developers.google.com/sheets/api/guides/formats
"""
function _format_number!(client::GoogleSheetsClient, sheet::Sheet, start_row_index::Integer, end_row_index::Integer, start_col_index::Integer, end_col_index::Integer, format_type::AbstractString, format_pattern::AbstractString)::Dict{Any,Any}
    body = Dict(
        "requests" => [
            Dict(
                "repeatCell" => Dict(
                    "range" => cellRange2D(sheet, start_row_index, end_row_index, start_col_index, end_col_index),

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

    return batch_update!(client, sheet.spreadsheet, body)
end


"""
Sets the background color.

See: https://developers.google.com/sheets/api/guides/formats
"""
function format_background_color!(client::GoogleSheetsClient, sheet::Sheet, start_row_index::Integer, end_row_index::Integer, start_col_index::Integer, end_col_index::Integer, color::Colorant)::Dict{Any,Any}
    body = Dict(
        "requests" => [
            Dict(
                "repeatCell" => Dict(
                    "range" => cellRange2D(sheet, start_row_index, end_row_index, start_col_index, end_col_index),

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

    return batch_update!(client, sheet.spreadsheet, body)
end