using ClosedXML.Excel;
using Microsoft.OpenApi.Models;
using OpenApi2Excel.Common;

namespace OpenApi2Excel.Builders.WorksheetPartsBuilders
{
    internal abstract class WorksheetPartBuilder(
        RowPointer actualRow,
        IXLWorksheet worksheet,
        OpenApiDocumentationOptions options)
    {
        protected OpenApiDocumentationOptions Options { get; } = options;
        protected RowPointer ActualRow { get; } = actualRow;
        protected IXLWorksheet Worksheet { get; } = worksheet;
        protected XLColor HeaderBackgroundColor
            => XLColor.LightGray;

        protected void AddEmptyRow()
            => ActualRow.MoveNext();

        protected CellBuilder Fill(int column)
            => new(Worksheet.Cell(ActualRow, column), Options);

        protected CellBuilder FillHeader(int column)
        {
            return Fill(column).WithBackground(HeaderBackgroundColor);
        }

        protected void AddPropertiesTreeForMediaTypes(IDictionary<string, OpenApiMediaType> mediaTypes, int attributesColumnIndex)
        {
            foreach (var mediaType in mediaTypes)
            {
                AddResponseFormat(mediaType.Key);
                AddResponseHeader();
                foreach (var property in mediaType.Value.Schema.Properties)
                {
                    AddProperty(property.Key, property.Value, 1, attributesColumnIndex);
                }
                AddEmptyRow();
            }
            AddEmptyRow();

            return;

            void AddResponseHeader()
            {
                var lastUsedColumn = FillSchemaDescriptionHeaderCells(attributesColumnIndex);
                ActualRow.MovePrev();
                FillHeaderBackground(1, lastUsedColumn);
                ActualRow.MoveNext(2);
            }

            void AddResponseFormat(string name)
            {
                Fill(1).WithText($"Format: {name}").WithBoldStyle();
                ActualRow.MoveNext();
            }
        }

        protected void AddProperty(string name, OpenApiSchema schema, int level, int attributesColumnIndex)
        {
            // current property
            AddPropertyRow(name, schema, level++);
            if (schema.Items != null)
            {
                if (schema.Items.Properties.Any())
                {
                    // array object subproperties
                    foreach (var property in schema.Items.Properties)
                    {
                        AddProperty(property.Key, property.Value, level, attributesColumnIndex);
                    }
                }
                else
                {
                    // if array contains simple type
                    AddProperty("element", schema.Items, level, attributesColumnIndex);
                }
            }

            // subproperties
            foreach (var property in schema.Properties)
            {
                AddProperty(property.Key, property.Value, level, attributesColumnIndex);
            }

            return;

            void AddPropertyRow(string propertyName, OpenApiSchema propertySchema, int propertyLevel)
            {
                FillHeaderBackground(1, propertyLevel - 1);
                FillCell(propertyLevel, propertyName);
                FillSchemaDescriptionCells(propertySchema, attributesColumnIndex);
                ActualRow.MoveNext();
            }
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

        protected int FillSchemaDescriptionHeaderCells(int attributesStartColumn)
        {
            var cellBuilder = Fill(1).WithText("Name").WithBoldStyle()
                .Next(attributesStartColumn - 1).WithText("Type").WithBoldStyle()
                .Next().WithText("Format").WithBoldStyle()
                .Next().WithText("Length").WithBoldStyle()
                .Next().WithText("Range").WithBoldStyle()
                .Next().WithText("Pattern").WithBoldStyle()
                .Next().WithText("Description").WithBoldStyle();

            var lastUsedColumn = cellBuilder.GetCell().Address.ColumnNumber;
            AddBottomBorder(1, lastUsedColumn);
            FillHeaderBackground(1, lastUsedColumn);
            return lastUsedColumn;
        }

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
