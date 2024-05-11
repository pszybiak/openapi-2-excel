using openapi2excel.core;
using Spectre.Console;
using Spectre.Console.Cli;
using System.ComponentModel;

namespace OpenApi2Excel.cli;

[Description("Generate Rest API specification in a MS Excel format")]
public class GenerateExcelCommand : Command<GenerateExcelCommand.GenerateExcelSettings>
{
   public class GenerateExcelSettings : CommandSettings
   {
      [Description("The path or URL to a YAML or JSON file with Rest API specification.")]
      [CommandArgument(0, "<INPUT_FILE>")]
      public string InputFile { get; init; } = null!;

      [Description("The path for output excel file.")]
      [CommandArgument(1, "<OUTPUT_FILE>")]
      public string OutputFile { get; init; } = null!;

      [Description("Run tool without logo.")]
      [CommandOption("-n|--no-logo")]
      public bool NoLogo { get; init; }

      internal FileInfo InputFileParsed { get; set; } = null!;
      internal FileInfo OutputFileParsed { get; set; } = null!;
      public override ValidationResult Validate()
      {
         var inputFilePath = InputFile.Trim();
         if (File.Exists(inputFilePath))
         {
            InputFileParsed = new FileInfo(inputFilePath);
         }
         else if (Uri.TryCreate(inputFilePath, UriKind.RelativeOrAbsolute, out var uri))
         {
            var inputFileTempPath = Path.Combine(Path.GetTempPath(), Path.GetRandomFileName() + ".xlsx");
            if (TryDownloadFileTaskAsync(uri, inputFileTempPath).GetAwaiter().GetResult())
            {
               InputFileParsed = new FileInfo(inputFileTempPath);
            }
            else
            {
               return ValidationResult.Error("Invalid input file path.");
            }
         }
         else
         {
            return ValidationResult.Error("Invalid input file path.");
         }

         var outputFilePath = OutputFile.Trim();
         if (!outputFilePath.EndsWith(".xlsx", StringComparison.CurrentCultureIgnoreCase))
         {
            outputFilePath += ".xlsx";
         }

         OutputFileParsed = new FileInfo(outputFilePath);

         return ValidationResult.Success();
      }

      private static async Task<bool> TryDownloadFileTaskAsync(Uri uri, string fileName)
      {
         try
         {
            var client = new HttpClient();
            await using var s = await client.GetStreamAsync(uri);
            await using var fs = new FileStream(fileName, FileMode.CreateNew);
            await s.CopyToAsync(fs);
            return true;
         }
         catch
         {
            return false;
         }
      }
   }

   public override int Execute(CommandContext context, GenerateExcelSettings settings)
   {
      if (!settings.NoLogo)
      {
         foreach (var renderable in CustomHelpProvider.GetHeaderText())
         {
            AnsiConsole.Write(renderable);
         }
      }

      try
      {
         OpenApiDocumentationGenerator
            .GenerateDocumentation(settings.InputFileParsed.FullName, settings.OutputFileParsed.FullName,
               new OpenApiDocumentationOptions())
            .ConfigureAwait(false).GetAwaiter().GetResult();

         AnsiConsole.MarkupLine($"Excel file saved to [green]{settings.OutputFileParsed.FullName}[/]");
      }
      catch (IOException exc)
      {
         AnsiConsole.MarkupLine($"[red]{exc.Message}[/]");
         return 1;
      }
      catch (Exception exc)
      {
         AnsiConsole.MarkupLine($"[red]An unexpected error occurred: {exc}[/]");
         return 1;
      }

      return 0;
   }
}