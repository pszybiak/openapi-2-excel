using ClosedXML.Excel;

namespace OpenApi2Excel.Builders.WorksheetPartsBuilders;

internal class CellBuilder(IXLCell cell, OpenApiDocumentationOptions option)
{
    public IXLCell GetCell() => cell;
    public int GetCellNumber() => cell.Address.ColumnNumber;

    public CellBuilder Next()
    {
        cell = cell.CellRight();
        return this;
    }

    public CellBuilder Next(int step)
    {
        cell = cell.CellRight(step);
        return this;
    }

    public CellBuilder Prev()
    {
        cell = cell.CellLeft();
        return this;
    }

    public CellBuilder Prev(int step)
    {
        cell = cell.CellLeft(step);
        return this;
    }

    public CellBuilder GoTo(int column)
    {
        var currentCell = GetCellNumber();
        if (currentCell == column)
            return this;
        return currentCell < column ? Next(column - currentCell) : Prev(currentCell - column);
    }

    public CellBuilder WithText(string? value)
        => With(c => c.Value = value);

    public CellBuilder WithBackground(XLColor color)
        => With(c => c.Style.Fill.BackgroundColor = color);

    public CellBuilder WithBackground(XLColor color, int endColumn)
    {
        while (GetCellNumber() <= endColumn)
            WithBackground(color).Next();
        return this;
    }

    public CellBuilder WithHorizontalAlignment(XLAlignmentHorizontalValues alignment)
        => With(c => c.Style.Alignment.SetHorizontal(alignment));

    public CellBuilder WithHorizontalAlignment(XLAlignmentHorizontalValues alignment, int endColumn)
    {
        while (GetCellNumber() <= endColumn)
            WithHorizontalAlignment(alignment).Next();
        return this;
    }

    public CellBuilder WithBottomBorder()
        => With(c => c.Style.Border.BottomBorder = XLBorderStyleValues.Medium);

    public CellBuilder WithBottomBorder(int endColumn)
    {
        while (GetCellNumber() <= endColumn)
            WithBottomBorder().Next();
        return this;
    }

    public CellBuilder WithBoldStyle()
        => With(c => c.Style.Font.Bold = true);

    private CellBuilder With(Action<IXLCell> action)
    {
        action(cell);
        return this;
    }
}
