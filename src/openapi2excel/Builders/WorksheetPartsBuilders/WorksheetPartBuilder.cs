using ClosedXML.Excel;
using Microsoft.OpenApi.Models;
using OpenApi2Excel.Common;

namespace OpenApi2Excel.Builders.WorksheetPartsBuilders
{
    internal abstract class WorksheetPartBuilder(RowPointer actualRow, IXLWorksheet worksheet, OpenApiDocumentationOptions options)
    {
        protected OpenApiDocumentationOptions Options { get; } = options;
        protected RowPointer ActualRow { get; } = actualRow;
        protected IXLWorksheet Worksheet { get; } = worksheet;
        protected XLColor HeaderBackgroundColor
            => XLColor.LightGray;


        protected void MoveToPrewRow()
            => ActualRow.MovePrev();

        protected void AddEmptyRow()
            => ActualRow.MoveNext();

        protected CellBuilder Fill(int column)
            => new(Worksheet.Cell(ActualRow, column), Options);

        protected CellBuilder FillHeader(int column)
        {
            return Fill(column).WithBackground(HeaderBackgroundColor);
        }

        protected void FillSchemaDescriptionCells(OpenApiSchema schema, int startColumn)
        {
            Fill(startColumn).WithText(schema.Type)
                .Next().WithText(schema.Format)
                .Next().WithText(schema.GetPropertyLengthDescription()).WithHorizontalAlignment(XLAlignmentHorizontalValues.Center)
                .Next().WithText(schema.GetPropertyRangeDescription()).WithHorizontalAlignment(XLAlignmentHorizontalValues.Center)
                .Next().WithText(schema.Pattern)
                .Next().WithText(schema.Description);
        }

        protected int FillSchemaDescriptionHeaderCells(int startColumn)
        {
            var cellBuilder = FillHeader(startColumn).WithText("Type")
                .Next().WithText("Format")
                .Next().WithText("Length")
                .Next().WithText("Range")
                .Next().WithText("Pattern")
                .Next().WithText("Description");
            return cellBuilder.GetCell().Address.ColumnNumber;
        }

        protected void FillHeaderCell(string? label, int columnIndex)
            => FillCell(columnIndex, label, HeaderBackgroundColor);

        protected void FillCell(int column, string? value, XLColor? backgoundColor = null, XLAlignmentHorizontalValues alignment = XLAlignmentHorizontalValues.Left)
        {
            var cellBuilder = Fill(column).WithText(value);
            if (backgoundColor is not null)
            {
                cellBuilder.WithBackground(backgoundColor);
            }
            cellBuilder.WithHorizontalAlignment(alignment);
        }

        protected void FillCell(int column, bool value, XLColor? backgoundColor = null)
        {
            var cellBuilder = Fill(column).WithText(Options.Language.Get(value));
            if (backgoundColor is not null)
            {
                cellBuilder.WithBackground(backgoundColor);
            }
        }

        protected void FillBackground(int startColumn, int endColumn, XLColor backgroundColor)
        {
            var cellBuilder = Fill(startColumn);
            for (var columnIndex = startColumn; columnIndex <= endColumn; columnIndex++)
            {
                cellBuilder.WithBackground(backgroundColor).Next();
            }
        }

        protected void FillHeaderBackground(int startColumn, int endColumn)
            => FillBackground(startColumn, endColumn, HeaderBackgroundColor);

        protected void AddBottomBorder(int startColumn, int endColumn)
        {
            var cellBuilder = Fill(startColumn);
            for (var columnIndex = startColumn; columnIndex <= endColumn; columnIndex++)
            {
                cellBuilder.WithBottomBorder().Next();
            }
        }
    }
}
