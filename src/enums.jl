
"""
Create an enum and export the enum type and all values.
"""
macro _exported_enum(name, args...)
    return esc(quote
        @enum($name, $(args...))
        export $name
        $([:(export $arg) for arg in args]...)
    end)
end


"""
Create and export a GoogleSheets enum.  The enum values are prefixed to avoid collisions.

Also creates a function which returns the GoogleSheets string value.
"""
macro _gsheet_enum(type_name, prefix, args...)
    values = [ Symbol(prefix,arg) for arg in args ]
    map = zip(values,String.(args))

    return esc(quote
        @_exported_enum($type_name, $(values...))
        function gsheet_string(x::$(type_name))::String
            d = Dict( 
                $([:($k => $v) for (k,v) in map]...)
            )

            return d[x]
        end
    end)
end


"""
An authorization scope for accessing Google resources.

AUTH_SCOPE_READONLY:  Allows read-only access to the user's sheets and their properties.
AUTH_SCOPE_READWRITE: Allows read/write access to the user's sheets and their properties.
"""
@_gsheet_enum AuthScope AUTH_SCOPE_ READONLY READWRITE


"""
A Google Sheets value type.

VALUE_TYPE_MIN: minimum value in a range
VALUE_TYPE_MAX: maximum value in a range
VALUE_TYPE_NUMBER: exact number
VALUE_TYPE_PERCENT: percentage of a range
VALUE_TYPE_PERCENTILE: percentile of a range
"""
@_gsheet_enum ValueType VALUE_TYPE_ MIN MAX NUMBER PERCENT PERCENTILE


"""
A Google Sheets condition type.

CONDITION_TYPE_NUMBER_GREATER: value must be greater than the condition value.
CONDITION_TYPE_NUMBER_GREATER_THAN_EQ: value must be greater than or equal to the condition value.
CONDITION_TYPE_NUMBER_LESS: value must be less than the condition value.
CONDITION_TYPE_NUMBER_LESS_THAN_EQ: value must be less than or equal to the condition value.
CONDITION_TYPE_NUMBER_EQ: value must be equal to the condition value.
CONDITION_TYPE_NUMBER_NOT_EQ: value must not be equal to the condition value.
CONDITION_TYPE_NUMBER_BETWEEN: value must be between the two condition values.
CONDITION_TYPE_NUMBER_NOT_BETWEEN: value must not be between the two condition values.
CONDITION_TYPE_TEXT_CONTAINS: value must contain the condition value.
CONDITION_TYPE_TEXT_NOT_CONTAINS: value must not contain the condition value.
CONDITION_TYPE_TEXT_STARTS_WITH: value must start with the condition value.
CONDITION_TYPE_TEXT_ENDS_WITH: value must end with the condition value.
CONDITION_TYPE_TEXT_EQ: value must equal the condition value.
CONDITION_TYPE_TEXT_NOT_EQ: value must not equal the condition value.
CONDITION_TYPE_TEXT_IS_EMAIL: value must be a valid email address.
CONDITION_TYPE_TEXT_IS_URL: value must be a valid URL.
CONDITION_TYPE_DATE_EQ: value must have the same date as the condition value.
CONDITION_TYPE_DATE_NOT_EQ: value must not have the same date as the condition value.
CONDITION_TYPE_DATE_BEFORE: value's date must be before the date of the condition value.
CONDITION_TYPE_DATE_AFTER: value's date must be before the date of the condition value.
CONDITION_TYPE_DATE_ON_OR_BEFORE: value's date must be on or before the date of the condition value.
CONDITION_TYPE_DATE_ON_OR_AFTER: value's date must be on or after the date of the condition value.
CONDITION_TYPE_DATE_BETWEEN: value's date must be between the dates of the condition values.
CONDITION_TYPE_DATE_NOT_BETWEEN: value's date must not be between the dates of the condition values.
CONDITION_TYPE_DATE_IS_VALID: value must be a date.
CONDITION_TYPE_ONE_OF_LIST: value must be present in the condition values.
CONDITION_TYPE_BLANK: value must be empty.
CONDITION_TYPE_NOT_BLANK: value must not be empty.
CONDITION_TYPE_CUSTOM_FORMULA: condition formula must evaluate to true.
CONDITION_TYPE_BOOLEAN: value must be TRUE/FALSE.
CONDITION_TYPE_ONE_OF_RANGE: value is present in the condition value's cell range.
"""
@_gsheet_enum ConditionType CONDITION_TYPE_ NUMBER_GREATER NUMBER_GREATER_THAN_EQ NUMBER_LESS NUMBER_LESS_THAN_EQ NUMBER_EQ NUMBER_NOT_EQ NUMBER_BETWEEN NUMBER_NOT_BETWEEN TEXT_CONTAINS TEXT_NOT_CONTAINS TEXT_STARTS_WITH TEXT_ENDS_WITH TEXT_EQ TEXT_NOT_EQ TEXT_IS_EMAIL TEXT_IS_URL DATE_EQ DATE_NOT_EQ DATE_BEFORE DATE_AFTER DATE_ON_OR_BEFORE DATE_ON_OR_AFTER DATE_BETWEEN DATE_NOT_BETWEEN DATE_IS_VALID ONE_OF_LIST BLANK NOT_BLANK CUSTOM_FORMULA BOOLEAN ONE_OF_RANGE


"""
A Google Sheets number format type.

NUMBER_FORMAT_TYPE_TEXT: text formatting, e.g. 1000.12
NUMBER_FORMAT_TYPE_NUMBER: number formatting, e.g. 1,000.12

NUMBER_FORMAT_TYPE_DATE: date formatting, e.g. 9/26/2008
NUMBER_FORMAT_TYPE_TIME: time formatting, e.g. 3:59:00 PM
NUMBER_FORMAT_TYPE_DATE_TIME: date+time formatting, e.g. 9/26/08 15:59:00

NUMBER_FORMAT_TYPE_SCIENTIFIC: scientific number formatting, e.g. 1.01E+03
"""
@_gsheet_enum NumberFormatType NUMBER_FORMAT_TYPE_ TEXT NUMBER PERCENT CURRENCY DATE TIME DATE_TIME SCIENTIFIC