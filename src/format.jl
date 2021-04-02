
export format!


"""
Formats a range of cells.
"""
function format!(client::GoogleSheetsClient, range::CellIndexRange2D, format::CellFormat)::Dict{Any,Any}
    body = Dict(
        "requests" => [
            Dict(
                "repeatCell" => Dict(
                    "range" => gsheet_json(range), 
                    "cell" => Dict(
                        "userEnteredFormat" => gsheet_json(format),
                    ),
                    # "fields" => "userEnteredFormat.numberFormat",
                    "fields" => "userEnteredFormat",
                ),
            ),
        ],
    )

    return batch_update!(client, range.sheet.spreadsheet, body)
end



