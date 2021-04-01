export format_number!, format_datetime!, format_background_color!


"""
Formats number values.

See: https://developers.google.com/sheets/api/guides/formats
"""
function format_number!(client::GoogleSheetsClient, range::CellIndexRange2D, format_pattern::AbstractString)::Dict{Any,Any}
    return _format_number!(client, range, "NUMBER", format_pattern)
end


"""
Formats date-time values.

See: https://developers.google.com/sheets/api/guides/formats
"""
function format_datetime!(client::GoogleSheetsClient, range::CellIndexRange2D, format_pattern::AbstractString)::Dict{Any,Any}
    return _format_number!(client, range, "DATE", format_pattern)
end


"""
Formats number values.

See: https://developers.google.com/sheets/api/guides/formats
"""
function _format_number!(client::GoogleSheetsClient, range::CellIndexRange2D, format_type::AbstractString, format_pattern::AbstractString)::Dict{Any,Any}
    body = Dict(
        "requests" => [
            Dict(
                "repeatCell" => Dict(
                    "range" => gsheet_json(range), 

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

    return batch_update!(client, range.sheet.spreadsheet, body)
end


"""
Sets the background color.

See: https://developers.google.com/sheets/api/guides/formats
"""
function format_background_color!(client::GoogleSheetsClient, range::CellIndexRange2D, color::Colorant)::Dict{Any,Any}
    body = Dict(
        "requests" => [
            Dict(
                "repeatCell" => Dict(
                    "range" => gsheet_json(range), 

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

    return batch_update!(client, range.sheet.spreadsheet, body)
end