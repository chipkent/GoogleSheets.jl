

"""
    gsheet_json(color::Colorant)

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
    gsheet_json(range::CellIndexRange1D, dim::AbstractString)

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
    gsheet_json(range::CellIndexRange2D)

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


"""
    gsheet_json(format::CellFormat)

Returns a dictionary of values to describe a Format.
"""
function gsheet_json(format::CellFormat)
    rst = Dict{String,Any}()

    number_format = Dict{String,Any}()
    if !isnothing(format.number_format_type)  number_format["type"] = gsheet_string(format.number_format_type) end
    if !isnothing(format.number_format_pattern)  number_format["pattern"] = format.number_format_pattern end
    if length(number_format) > 0  rst["numberFormat"] = number_format end

    text_format = Dict{String,Any}()
    if !isnothing(format.text_bold)  text_format["bold"] = true end
    if !isnothing(format.text_italic)  text_format["italic"] = true end
    if !isnothing(format.text_color)  text_format["foreground_color"] = gsheet_json(format.text_color) end
    if !isnothing(format.text_font_size)  text_format["fontSize"] = format.text_font_size end
    if length(text_format) > 0  rst["textFormat"] = text_format end

    if !isnothing(format.background_color)  rst["backgroundColor"] = gsheet_json(format.background_color) end

    return rst
end

