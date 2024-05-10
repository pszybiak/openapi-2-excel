using ClosedXML.Excel;
using Microsoft.OpenApi.Models;
using openapi2excel.core.Builders.WorksheetPartsBuilders;
using openapi2excel.core.Builders.WorksheetPartsBuilders.Common;
using openapi2excel.core.Common;

namespace openapi2excel.core.Builders;

internal class OperationWorksheetBuilder(IXLWorkbook workbook, OpenApiDocumentationOptions options)
   : WorksheetBuilder(options)
{
   private readonly RowPointer _actualRowPointer = new(1);
   private IXLWorksheet _worksheet = null!;
   private int _attributesColumnsStartIndex;

   public IXLWorksheet Build(string path, OpenApiPathItem pathItem, OperationType operationType,
      OpenApiOperation operation)
   {
      var worksheet = GetWorksheetName(path, operation);
      CreateNewWorksheet(worksheet);
      _actualRowPointer.GoTo(1);

      _attributesColumnsStartIndex = MaxPropertiesTreeLevel.Calculate(operation);
      AdjustColumnsWidthToRequestTreeLevel();

      AddHomePageLink();
      AddOperationInfos(path, pathItem, operationType, operation);
      AddRequestParameters(operation);
      AddRequestBody(operation);
      AddResponseBody(operation);
      AdjustLastNamesColumnToContents();
      AdjustDescriptionColumnToContents();

      return _worksheet;
   }

   private static string GetWorksheetName(string path, OpenApiOperation operation)
   {
      if (!string.IsNullOrEmpty(operation.OperationId))
      {
         return operation.OperationId;
      }
      var startIndex = path.LastIndexOf("/", StringComparison.Ordinal) + 1;
      return path[startIndex..];
   }

   private void CreateNewWorksheet(string operation)
   {
      _worksheet = workbook.Worksheets.Add(operation);
      _worksheet.Style.Font.FontSize = 10;
      _worksheet.Style.Font.FontName = "Arial";
      _worksheet.Outline.SummaryHLocation = XLOutlineSummaryHLocation.Left;
      _worksheet.Outline.SummaryVLocation = XLOutlineSummaryVLocation.Top;
   }

   private void AdjustColumnsWidthToRequestTreeLevel()
   {
      for (var columnIndex = 1; columnIndex < _attributesColumnsStartIndex - 1; columnIndex++)
      {
         _worksheet.Column(columnIndex).Width = 1.8;
      }
   }

   private void AdjustLastNamesColumnToContents()
   {
      if (_attributesColumnsStartIndex > 1)
      {
         _worksheet.Column(_attributesColumnsStartIndex - 1).AdjustToContents();
      }
   }

   private void AdjustDescriptionColumnToContents()
   {
      _worksheet.LastColumnUsed().AdjustToContents();
   }

   private void AddOperationInfos(string path, OpenApiPathItem pathItem, OperationType operationType,
      OpenApiOperation operation) =>
      new OperationInfoBuilder(_actualRowPointer, _attributesColumnsStartIndex, _worksheet, Options)
         .AddOperationInfoSection(path, pathItem, operationType, operation);

   private void AddRequestParameters(OpenApiOperation operation) =>
      new RequestParametersBuilder(_actualRowPointer, _attributesColumnsStartIndex, _worksheet, Options)
         .AddRequestParametersPart(operation);

   private void AddRequestBody(OpenApiOperation operation) =>
      new RequestBodyBuilder(_actualRowPointer, _attributesColumnsStartIndex, _worksheet, Options)
         .AddRequestBodyPart(operation);

   private void AddResponseBody(OpenApiOperation operation) =>
      new ResponseBodyBuilder(_actualRowPointer, _attributesColumnsStartIndex, _worksheet, Options)
         .AddResponseBodyPart(operation);

   private void AddHomePageLink() => new HomePageLinkBuilder(_actualRowPointer, _worksheet, Options)
      .AddHomePageLinkSection();
}