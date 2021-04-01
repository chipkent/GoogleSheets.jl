

"""
Returns a dictionary of values to describe a color.
"""
function gsheet_json(color::Colorant)
    c = convert(RGBA, color)
    return Dict(
            "red" => red(c),
            "green" => green(c),
            "blue" => blue(c),
            "alpha" => alpha(c),
        )
end


"""
Returns a dictionary of values to describe a CellIndexRange1D.
"""
function gsheet_json(range::CellIndexRange1D, dim::AbstractString)
    return Dict(
        "sheetId" => range.sheet.id,
        "dimension" => dim,
        "startIndex" => range.start_index,
        "endIndex" => range.end_index,
    )
end


"""
Returns a dictionary of values to describe a CellIndexRange2D.
"""
function gsheet_json(range::CellIndexRange2D)
    return Dict(
            "sheetId" => range.sheet.id,
            "startRowIndex" => range.start_row_index,
            "endRowIndex" => range.end_row_index+1,
            "startColumnIndex" => range.start_col_index,
            "endColumnIndex" => range.end_col_index+1,
        )
end
