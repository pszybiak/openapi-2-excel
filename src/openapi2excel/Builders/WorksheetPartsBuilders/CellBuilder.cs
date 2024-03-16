using ClosedXML.Excel;

namespace OpenApi2Excel.Builders.WorksheetPartsBuilders;

internal class CellBuilder(IXLCell cell, OpenApiDocumentationOptions option)
{
    public CellBuilder Next()
    {
        cell = cell.CellRight();
        return this;
    }

    public CellBuilder Prev()
    {
        cell = cell.CellLeft();
        return this;
    }

    public CellBuilder WithText(string value)
        => With(c => c.Value = value);

    public CellBuilder WithBackground(XLColor color)
        => With(c => c.Style.Fill.BackgroundColor = color);

    public CellBuilder WithHorizontalAlignment(XLAlignmentHorizontalValues alignment)
        => With(c => c.Style.Alignment.SetHorizontal(alignment));

    public CellBuilder WithBottomBorder()
        => With(c => c.Style.Border.BottomBorder = XLBorderStyleValues.Medium);

    private CellBuilder With(Action<IXLCell> action)
    {
        action(cell);
        return this;
    }
}