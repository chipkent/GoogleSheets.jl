export format_color_scale!


"""
Sets color scale formatting.
"""
function format_color_scale!(client::GoogleSheetsClient, spreadsheet::Spreadsheet, title::AbstractString, 
    start_row_index::Integer, end_row_index::Integer, start_col_index::Integer, end_col_index::Integer; 
    min_color::Colorant=colorant"salmon", min_value_type::ValueType=VALUE_TYPE_MIN, min_value::Union{Nothing,Number}=nothing, 
    mid_color::Union{Nothing,Colorant}=nothing, mid_value_type::Union{Nothing,ValueType}=nothing, mid_value::Union{Nothing,Number}=nothing, 
    max_color::Colorant=colorant"springgreen", max_value_type::ValueType=VALUE_TYPE_MAX, max_value::Union{Nothing,Number}=nothing)::Dict{Any,Any}

    properties = meta(client, spreadsheet, title)
    return format_color_scale!(client, spreadsheet, properties["sheetId"], start_row_index, end_row_index, start_col_index, end_col_index; 
        min_color=min_color, min_value_type=min_value_type, min_value=min_value, mid_color=mid_color, mid_value_type=mid_value_type, mid_value=mid_value, 
        max_color=max_color, max_value_type=max_value_type, max_value=max_value)
end


"""
Sets color scale formatting.
"""
function format_color_scale!(client::GoogleSheetsClient, spreadsheet::Spreadsheet, sheet_id::Int64, 
    start_row_index::Integer, end_row_index::Integer, start_col_index::Integer, end_col_index::Integer; 
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

    function value_type(x)
        d = Dict(
            VALUE_TYPE_MIN => "MIN",
            VALUE_TYPE_MAX => "MAX",
            VALUE_TYPE_NUMBER => "NUMBER", 
            VALUE_TYPE_PERCENT => "PERCENT",
            VALUE_TYPE_PERCENTILE => "PERCENTILE",
        )

        return isnothing(x) ? nothing : d[x]
    end

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
                        "ranges" => [cellRange2D(sheet_id, start_row_index, end_row_index, start_col_index, end_col_index)],
                        "gradientRule" => gradientRule(),
                    ),
                ),
            ),
        ],
    )

    return batch_update!(client, spreadsheet, body)
end

# ***** TODO: implement below


# """
# Sets color scale formatting.
# """
# function format_color_scale!(client::GoogleSheetsClient, spreadsheet::Spreadsheet, title::AbstractString, 
#     start_row_index::Integer, end_row_index::Integer, start_col_index::Integer, end_col_index::Integer; 
#     min_color::Colorant=colorant"salmon", min_value_type::ValueType=VALUE_TYPE_MIN, min_value::Union{Nothing,Number}=nothing, 
#     mid_color::Union{Nothing,Colorant}=nothing, mid_value_type::Union{Nothing,ValueType}=nothing, mid_value::Union{Nothing,Number}=nothing, 
#     max_color::Colorant=colorant"springgreen", max_value_type::ValueType=VALUE_TYPE_MAX, max_value::Union{Nothing,Number}=nothing)::Dict{Any,Any}

#     properties = meta(client, spreadsheet, title)
#     return format_color_scale!(client, spreadsheet, properties["sheetId"], start_row_index, end_row_index, start_col_index, end_col_index; 
#         min_color=min_color, min_value_type=min_value_type, min_value=min_value, mid_color=mid_color, mid_value_type=mid_value_type, mid_value=mid_value, 
#         max_color=max_color, max_value_type=max_value_type, max_value=max_value)
# end


# """
# Sets color scale formatting.
# """
# function format_color_scale!(client::GoogleSheetsClient, spreadsheet::Spreadsheet, sheet_id::Int64, 
#     start_row_index::Integer, end_row_index::Integer, start_col_index::Integer, end_col_index::Integer, color::Colorant, values::Array; 
#     *****
#     min_color::Colorant=colorant"salmon", min_value_type::ValueType=VALUE_TYPE_MIN, min_value::Union{Nothing,Number}=nothing, 
#     mid_color::Union{Nothing,Colorant}=nothing, mid_value_type::Union{Nothing,ValueType}=nothing, mid_value::Union{Nothing,Number}=nothing, 
#     max_color::Colorant=colorant"springgreen", max_value_type::ValueType=VALUE_TYPE_MAX, max_value::Union{Nothing,Number}=nothing)::Dict{Any,Any}
 
#     # function point(color, type, value)
#     #     rst = Dict{Any,Any}(
#     #             "color" => gsheet_json(color),
#     #     )

#     #     if !isnothing(type)
#     #         rst["type"] = type
#     #     end

#     #     if !isnothing(value)
#     #         rst["value"] = "$value"
#     #     end

#     #     return rst
#     # end

#     #TODO: remove ******************************************
#     function value_type(x)
#         d = Dict(
#             VALUE_TYPE_MIN => "MIN",
#             VALUE_TYPE_MAX => "MAX",
#             VALUE_TYPE_NUMBER => "NUMBER", 
#             VALUE_TYPE_PERCENT => "PERCENT",
#             VALUE_TYPE_PERCENTILE => "PERCENTILE",
#         )

#         return isnothing(x) ? nothing : d[x]
#     end

#     body = Dict(
#         "requests" => [
#             Dict(
#                 "addConditionalFormatRule" => Dict(
#                     "rule" => Dict(
#                         "ranges" => [cellRange2D(sheet_id, start_row_index, end_row_index, start_col_index, end_col_index)],
#                         "booleanRule" => Dict(
#                             "condition" => Dict(
#                                 "type" => "NUMBER_LESS_THAN_EQ"***,
#                                 "values" => [ Dict("userEnteredValue" => "$value") for value in values ],
#                             ),
#                             "format" => Dict(
#                                 "backgroundColor" => gsheet_json(color),
#                             ),
#                         ),
#                     ),
#                 ),
#             ),
#         ],
#     )

#     return batch_update!(client, spreadsheet, body)
# end