
export GoogleSheetsClient, Spreadsheet, Sheet, CellRange, CellRanges, CellRangeValues, UpdateSummary, CellIndexRange1D, CellIndexRange2D


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
A sheet in a spreadsheet.
"""
Sheet(client::GoogleSheetsClient, spreadsheet::Spreadsheet, id::Int64) = Sheet(spreadsheet, id, meta(client, spreadsheet, id)["title"])


"""
A sheet in a spreadsheet.
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

