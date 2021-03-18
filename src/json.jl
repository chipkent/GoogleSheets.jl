

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