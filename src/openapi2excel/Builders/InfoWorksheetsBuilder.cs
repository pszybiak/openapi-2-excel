using ClosedXML.Excel;
using Microsoft.OpenApi.Models;

namespace openapi2excel.Builders;

internal class InfoWorksheetBuilder : WorksheetBuilder
{
    private readonly OpenApiDocument _readResultOpenApiDocument;
    public static string Name => "Info";
    private readonly IXLWorksheet _infoWorksheet;
    private int _actualRowIndex = 1;

    public InfoWorksheetBuilder(IXLWorkbook workbook, OpenApiDocument readResultOpenApiDocument, OpenApiDocumentationOptions options)
    : base(options)
    {
        _readResultOpenApiDocument = readResultOpenApiDocument;
        _infoWorksheet = workbook.Worksheets.Add(Name);
        _infoWorksheet.Column(1).Width = 11;

        AddVersion();
        AddTitle();
        AddDescription();
        AddEmptyRow();
    }

    public void AddLink(OperationType operation, string path, IXLWorksheet worksheet)
    {
        _infoWorksheet.Cell(_actualRowIndex, 1).Value = operation.ToString().ToUpper();
        _infoWorksheet.Cell(_actualRowIndex, 2).Value = path;

        _infoWorksheet.Cell(_actualRowIndex, 1).SetHyperlink(new XLHyperlink($"'{worksheet.Name}'!A1"));
        _infoWorksheet.Cell(_actualRowIndex, 2).SetHyperlink(new XLHyperlink($"'{worksheet.Name}'!A1"));
        _actualRowIndex++;
    }

    private void AddVersion()
    {
        if (string.IsNullOrEmpty(_readResultOpenApiDocument.Info.Version))
        {
            return;
        }

        FillInfo("Version", _readResultOpenApiDocument.Info.Version);
    }

    private void AddDescription()
    {
        if (string.IsNullOrEmpty(_readResultOpenApiDocument.Info.Description))
        {
            return;
        }
        FillInfo("Description", _readResultOpenApiDocument.Info.Description, true);
    }

    private void AddTitle()
    {
        if (string.IsNullOrEmpty(_readResultOpenApiDocument.Info.Title))
        {
            return;
        }
        FillInfo("Title", _readResultOpenApiDocument.Info.Title);
    }

    private void AddEmptyRow()
    {
        _actualRowIndex++;
    }

    private void FillInfo(string name, string value, bool multipleRowText = false)
    {
        _infoWorksheet.Cell(_actualRowIndex, 1).Value = name;
        if (multipleRowText)
        {
            var splitValues = value.Split('\n', '\r', StringSplitOptions.RemoveEmptyEntries | StringSplitOptions.TrimEntries);
            foreach (var splitValue in splitValues)
            {
                _infoWorksheet.Cell(_actualRowIndex++, 2).Value = splitValue;
            }
        }
        else
        {
            _infoWorksheet.Cell(_actualRowIndex++, 2).Value = value;
        }
    }
}