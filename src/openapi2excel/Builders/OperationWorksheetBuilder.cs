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
      var worksheet = GetWorksheetName(path, operation, operationType);

      CreateNewWorksheet(worksheet);
      _actualRowPointer.GoTo(1);

      _attributesColumnsStartIndex = MaxPropertiesTreeLevel.Calculate(operation, Options.MaxDepth);
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

   private string GetWorksheetName(string path, OpenApiOperation operation, OperationType operationType)
   {
      var maxLength = 28;
      var name = "";
      if (!string.IsNullOrEmpty(operation.OperationId))
      {
         // take worksheet name from OperationId
         name = operation.OperationId;
      }
      else
      {
         // generate worksheet name based on operationType and path
         var pathName = path.Replace("/", "-");
         name = operationType.ToString().ToUpper() + "_" + pathName[1..];
      }

      // check if the name is not too long
      if (name.Length > maxLength)
      {
         name = name[..maxLength];
      }

      // check if the name is unique
      var nr = 2;
      var tmpName = name;
      while (workbook.Worksheets.Any(s => s.Name.Equals(tmpName, StringComparison.CurrentCultureIgnoreCase)))
      {
         tmpName = name[..maxLength] + "_" + nr++;
      }
      return tmpName;
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