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
        protected XLColor HeaderBackgroundColor => XLColor.LightGray;

        protected void MoveToNextRow()
            => ActualRow.MoveNext();

        protected void AddEmptyRow()
            => MoveToNextRow();

        protected CellBuilder Fill(int column)
        {
            return new CellBuilder(Worksheet.Cell(ActualRow, column), Options);
        }

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
            FillHeader(startColumn++).WithText("Type");
            FillHeader(startColumn++).WithText("Format");
            FillHeader(startColumn++).WithText("Length");
            FillHeader(startColumn++).WithText("Range");
            FillHeader(startColumn++).WithText("Pattern");
            FillHeader(startColumn).WithText("Description");
            return startColumn;
        }

        protected void FillHeaderCell(string? label, int columnIndex)
            => FillCell(columnIndex, label, HeaderBackgroundColor);

        protected void FillCell(int column, string? value, XLColor? backgoundColor = null, XLAlignmentHorizontalValues alignment = XLAlignmentHorizontalValues.Left)
        {
            Worksheet.Cell(ActualRow, column).Value = value;
            if (backgoundColor is not null)
            {
                Worksheet.Cell(ActualRow, column).Style.Fill.BackgroundColor = backgoundColor;
            }
            Worksheet.Cell(ActualRow, column).Style.Alignment.SetHorizontal(alignment);
        }

        protected void FillCell(int column, string value, XLAlignmentHorizontalValues alignment)
        {
            FillCell(column, value, null, alignment);
        }

        protected void FillCell(int column, bool value, XLColor? backgoundColor = null)
        {
            Worksheet.Cell(ActualRow, column).Value = Options.Language.Get(value);
            if (backgoundColor is not null)
            {
                Worksheet.Cell(ActualRow, column).Style.Fill.BackgroundColor = backgoundColor;
            }
        }

        protected void FillBackground(int startColumn, int endColumn, XLColor backgroundColor)
        {
            for (var columnIndex = startColumn; columnIndex <= endColumn; columnIndex++)
            {
                FillBackground(columnIndex, backgroundColor);
            }
        }

        protected void FillBackground(int column, XLColor backgroundColor)
        {
            Worksheet.Cell(ActualRow, column).Style.Fill.BackgroundColor = backgroundColor;
        }

        protected void FillHeaderBackground(int startColumn, int endColumn)
        {
            FillBackground(startColumn, endColumn, HeaderBackgroundColor);
        }

        protected void FillHeaderBackground(int column)
        {
            FillBackground(column, HeaderBackgroundColor);
        }

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
