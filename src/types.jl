
export GoogleSheetsClient, Spreadsheet, CellRange, CellRanges, CellRangeValues, UpdateSummary


"""
A Google Sheets client.
"""
struct GoogleSheetsClient
    client
end


"""
A spreadsheet.
"""
struct Spreadsheet
    """Spreadsheet unique identifier."""
    id::AbstractString
end


"""
A range of cells within a spreadsheet.
"""
struct CellRange
    """Spreadsheet containing the cells."""
    spreadsheet::Spreadsheet

    """Range of cells."""
    range::AbstractString
end


"""
Multiple ranges of cells within a spreadsheet.
"""
struct CellRanges{T<:AbstractString}
    """Spreadsheet containing the cells."""
    spreadsheet::Spreadsheet

    """Ranges of cells."""
    ranges::Array{T,1}
end


"""
A range of cell values within a spreadsheet.
"""
struct CellRangeValues
    """Range of cells within a spreadsheet."""
    range::CellRange

    """Values of cells within a spreadsheet."""
    values::Union{Nothing,Array{String,2}}

    """Major dimension of the cell values."""
    major_dimension::AbstractString
end


"""
Summary of updated updated cells.
"""
struct UpdateSummary
    """Range of cells within a spreadsheet."""
    range::CellRange

    """Number of updated columns."""
    updated_columns::Int64

    """Number of updated rows."""
    updated_rows::Int64

    """Number of updated cells."""
    updated_cells::Int64
end