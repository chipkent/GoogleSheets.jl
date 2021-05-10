
export format!, format_conditional!, format_color_gradient!


"""
    format!(client::GoogleSheetsClient, range::CellIndexRange2D, format::CellFormat)::Dict{Any,Any}

Formats a range of cells.

# Arguments
- `client::GoogleSheetsClient`: client.
- `range::CellIndexRange2D`: cell index range.
- `format::CellFormat`: cell format.
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


"""
    format_conditional!(client::GoogleSheetsClient, range::CellIndexRange2D, format::CellFormat, 
        condition_type::ConditionType, values...)::Dict{Any,Any}

Formats a range of cells if a condition is met.

# Arguments
- `client::GoogleSheetsClient`: client.
- `range::CellIndexRange2D`: cell index range.
- `format::CellFormat`: cell format.
- `condition_type::ConditionType`: type of condition.
- `values...`: values for the condition.
"""
function format_conditional!(client::GoogleSheetsClient, range::CellIndexRange2D, format::CellFormat, 
    condition_type::ConditionType, values...)::Dict{Any,Any}
 
    body = Dict(
        "requests" => [
            Dict(
                "addConditionalFormatRule" => Dict(
                    "rule" => Dict(
                        "ranges" => [gsheet_json(range)],
                        "booleanRule" => Dict(
                            "condition" => Dict(
                                "type" => gsheet_string(condition_type),
                                "values" => [ Dict("userEnteredValue" => "$value") for value in values ],
                            ),
                            "format" => gsheet_json(format),
                        ),
                    ),
                ),
            ),
        ],
    )

    return batch_update!(client, range.sheet.spreadsheet, body)
end


"""
    format_color_gradient!(client::GoogleSheetsClient, range::CellIndexRange2D; 
        min_color::Colorant=colorant"salmon", min_value_type::ValueType=VALUE_TYPE_MIN, min_value::Union{Nothing,Number}=nothing, 
        mid_color::Union{Nothing,Colorant}=nothing, mid_value_type::Union{Nothing,ValueType}=nothing, mid_value::Union{Nothing,Number}=nothing, 
        max_color::Colorant=colorant"springgreen", max_value_type::ValueType=VALUE_TYPE_MAX, max_value::Union{Nothing,Number}=nothing)::Dict{Any,Any}

Sets color gradient formatting.

# Arguments
- `client::GoogleSheetsClient`: client.
- `range::CellIndexRange2D`: cell index range.

- `min_color::Colorant=colorant"salmon"`: color for the minimum value.
- `min_value_type::ValueType=VALUE_TYPE_MIN`: minimum value type.
- `min_value::Union{Nothing,Number}=nothing`: minimum value.
- `mid_color::Union{Nothing,Colorant}=nothing`: color for the mid value.
- `mid_value_type::Union{Nothing,ValueType}=nothing`: mid value type.
- `mid_value::Union{Nothing,Number}=nothing`: mid value.
- `max_color::Colorant=colorant"springgreen"`: color for the maximum value.
- `max_value_type::ValueType=VALUE_TYPE_MAX`: maximum value type.
- `max_value::Union{Nothing,Number}=nothing`: maximum value.
"""
function format_color_gradient!(client::GoogleSheetsClient, range::CellIndexRange2D; 
    min_color::Colorant=colorant"salmon", min_value_type::ValueType=VALUE_TYPE_MIN, min_value::Union{Nothing,Number}=nothing, 
    mid_color::Union{Nothing,Colorant}=nothing, mid_value_type::Union{Nothing,ValueType}=nothing, mid_value::Union{Nothing,Number}=nothing, 
    max_color::Colorant=colorant"springgreen", max_value_type::ValueType=VALUE_TYPE_MAX, max_value::Union{Nothing,Number}=nothing)::Dict{Any,Any}
 
    function point(color, type, value)
        rst = Dict{Any,Any}(
                "color" => gsheet_json(color),
        )

        if !isnothing(type)
            rst["type"] = type
        end

        if !isnothing(value)
            rst["value"] = "$value"
        end

        return rst
    end

    value_type(x) =return isnothing(x) ? nothing : gsheet_string(x)

    function gradientRule()
        value_type_min = value_type(min_value_type)
        value_type_max = value_type(max_value_type)
        value_type_mid = value_type(mid_value_type)

        rst = Dict(
                "minpoint" => point(min_color, value_type_min, min_value),
                "maxpoint" => point(max_color, value_type_max, max_value),
        )

        if !isnothing(mid_color)
            rst["midpoint"] = point(mid_color, value_type_mid, mid_value)
        end

        return rst
    end

    body = Dict(
        "requests" => [
            Dict(
                "addConditionalFormatRule" => Dict(
                    "rule" => Dict(
                        "ranges" => [gsheet_json(range)],
                        "gradientRule" => gradientRule(),
                    ),
                ),
            ),
        ],
    )

    return batch_update!(client, range.sheet.spreadsheet, body)
end
