
export GoogleSheetsClient, Spreadsheet, Sheet, CellRange, CellRanges, CellRangeValues, UpdateSummary, CellIndexRange1D, CellIndexRange2D, CellFormat


"""
A Google Sheets client.
"""
struct GoogleSheetsClient
    """ Client Python object. """
    client
    
    """ Rate limiter for API read calls. """
    rate_limiter_read::AbstractRateLimiter

    """ Rate limiter for API write calls. """
    rate_limiter_write::AbstractRateLimiter
end


"""
A spreadsheet.
"""
struct Spreadsheet
    """ Spreadsheet unique identifier. """
    id::AbstractString
end


"""
A sheet in a spreadsheet.
"""
struct Sheet
    """ Spreadsheet """
    spreadsheet::Spreadsheet

    """ Sheet unique identifier."""
    id::Int64 

    """ Sheet title. """
    title::AbstractString 
end


"""
    Sheet(client::GoogleSheetsClient, spreadsheet::Spreadsheet, id::Int64)

A sheet in a spreadsheet.

# Arguments
- `client::GoogleSheetsClient`: client
- `spreadsheet::Spreadsheet`: spreadsheet
- `id::Int64`: sheet id
"""
Sheet(client::GoogleSheetsClient, spreadsheet::Spreadsheet, id::Int64) = Sheet(spreadsheet, id, meta(client, spreadsheet, id)["title"])


"""
    Sheet(client::GoogleSheetsClient, spreadsheet::Spreadsheet, title::AbstractString)

A sheet in a spreadsheet.

# Arguments
- `client::GoogleSheetsClient`: client
- `spreadsheet::Spreadsheet`: spreadsheet
- `title::AbstractString`: sheet title
"""
Sheet(client::GoogleSheetsClient, spreadsheet::Spreadsheet, title::AbstractString) = Sheet(spreadsheet, meta(client, spreadsheet, title)["sheetId"], title)


"""
A range of cells within a spreadsheet.
"""
struct CellRange
    """ Spreadsheet containing the cells. """
    spreadsheet::Spreadsheet

    """ Range of cells. """
    range::AbstractString
end


"""
Multiple ranges of cells within a spreadsheet.
"""
struct CellRanges{T<:AbstractString}
    """ Spreadsheet containing the cells. """
    spreadsheet::Spreadsheet

    """ Ranges of cells. """
    ranges::Array{T,1}
end


"""
A range of cell values within a spreadsheet.
"""
struct CellRangeValues
    """ Range of cells within a spreadsheet. """
    range::CellRange

    """ Values of cells within a spreadsheet. """
    values::Union{Nothing,Array{String,2}}

    """ Major dimension of the cell values. """
    major_dimension::AbstractString
end


"""
Summary of updated updated cells.
"""
struct UpdateSummary
    """ Range of cells within a spreadsheet. """
    range::CellRange

    """ Number of updated columns. """
    updated_columns::Int64

    """ Number of updated rows. """
    updated_rows::Int64

    """ Number of updated cells. """
    updated_cells::Int64
end


"""
1D Range of cell values.
"""
struct CellIndexRange1D
    """ Sheet containing the cells. """
    sheet::Sheet

    """ Start index of the cells. """
    start_index::Integer

    """ End index of the cells. """
    end_index::Integer
end


"""
2D Range of cell values.
"""
struct CellIndexRange2D
    """ Sheet containing the cells. """
    sheet::Sheet

    """ Start row index of the cells. """
    start_row_index::Integer

    """ End row index of the cells. """
    end_row_index::Integer

    """ Start column index of the cells. """
    start_col_index::Integer
 
    """ End column index of the cells. """
    end_col_index::Integer
end


"""
Formatting for a cell.
"""
struct CellFormat
    """ Background color. """
    background_color::Union{Nothing,Colorant}

    """ Number format type. """
    number_format_type::Union{Nothing,NumberFormatType}

    """ 
    Number format pattern. 
    See: https://developers.google.com/sheets/api/guides/formats
    """
    number_format_pattern::Union{Nothing,AbstractString}

    """ Should the text be formatted in italics. """
    text_italic::Union{Nothing,Bool}

    """ Should the text be bold. """
    text_bold::Union{Nothing,Bool}

    """ Should the text be strikethrough. """
    text_strikethrough::Union{Nothing,Bool}

    """ Text color. """
    text_color::Union{Nothing,Colorant} 

    """ Text font size. """
    text_font_size::Union{Nothing,Real}

    """
    Formatting for a cell.

    # Arguments
    - `background_color=nothing`: background color
    - `number_format_type=nothing`: number format
    - `number_format_pattern=nothing`: number format pattern
    - `text_italic=nothing`: is text italics
    - `text_bold=nothing`: is text bold
    - `text_strikethrough=nothing`: is text strikethrough
    - `text_color=nothing`: text color
    - `text_font_size=nothing`: text_font_size
    """
    CellFormat(;
        background_color=nothing,
        number_format_type=nothing,
        number_format_pattern=nothing,
        text_italic=nothing,
        text_bold=nothing,
        text_strikethrough=nothing,
        text_color=nothing,
        text_font_size=nothing
        ) = new(background_color, number_format_type, number_format_pattern, text_italic, text_bold, text_strikethrough, text_color, text_font_size)
end