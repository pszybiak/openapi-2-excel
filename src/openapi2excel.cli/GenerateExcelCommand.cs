using openapi2excel.core;
using openapi2excel.core.Lang;
using openapi2excel.core.Sanitization;
using Spectre.Console;
using Spectre.Console.Cli;
using System.ComponentModel;
using Path = System.IO.Path;

namespace OpenApi2Excel.cli;

[Description("Generate Rest API specification in a MS Excel format")]
public class GenerateExcelCommand : AsyncCommand<GenerateExcelCommand.GenerateExcelSettings>
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

      [Description("Language of the generated document: a language code or a path to a json file with a translation.")]
      [CommandOption("-l|--lang")]
      public string? Language { get; init; }

      [Description("Maximum depth level for documenting object hierarchies (defaults to 10).")]
      [CommandOption("-d|--depth")]
      public int Depth { get; init; } = 10;

      [Description("Run tool with debug mode.")]
      [CommandOption("-g|--debug")]
      public bool Debug { get; init; }

      [Description("Reject the input file instead of correcting its specification errors.")]
      [CommandOption("--no-sanitize")]
      public bool NoSanitize { get; init; }

      internal FileInfo InputFileParsed { get; set; } = null!;
      internal FileInfo OutputFileParsed { get; set; } = null!;
      internal TranslationResult TranslationParsed { get; set; } = null!;

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

         try
         {
            TranslationParsed = TranslationsSelector.Load(Language);
         }
         catch (InvalidLanguageException exception)
         {
            return ValidationResult.Error(exception.Message);
         }

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

   protected override async Task<int> ExecuteAsync(CommandContext context, GenerateExcelSettings settings,
      CancellationToken cancellationToken)
   {
      if (!settings.NoLogo)
      {
         foreach (var renderable in CustomHelpProvider.GetHeaderText())
         {
            AnsiConsole.Write(renderable);
         }
      }

      WriteTranslationReport(settings);

      try
      {
         var options = new OpenApiDocumentationOptions
         {
            MaxDepth = settings.Depth,
            Translation = settings.TranslationParsed.Translation,
            SanitizeDocument = !settings.NoSanitize,
            OnDocumentSanitized = WriteSanitizationReport
         };

         await OpenApiDocumentationGenerator
            .GenerateDocumentation(settings.InputFileParsed.FullName, settings.OutputFileParsed.FullName, options)
            .ConfigureAwait(false);

         AnsiConsole.MarkupLine($"Excel file saved to [green]{settings.OutputFileParsed.FullName.EscapeMarkup()}[/]");
      }
      catch (IOException exc)
      {
         AnsiConsole.MarkupLine(settings.Debug ? $"[red]{exc.ToString().EscapeMarkup()}[/]" : $"[red]{exc.Message.EscapeMarkup()}[/]");

         return 1;
      }
      catch (Exception exc)
      {
         AnsiConsole.MarkupLine(settings.Debug
            ? $"[red]An unexpected error occurred: {exc.ToString().EscapeMarkup()}[/]"
            : $"[red]An unexpected error occurred: {exc.Message.EscapeMarkup()}[/]");

         return 1;
      }

      return 0;
   }

   private static void WriteTranslationReport(GenerateExcelSettings settings)
   {
      if (settings.TranslationParsed.IsComplete)
      {
         return;
      }

      var missing = settings.TranslationParsed.MissingLabels;
      AnsiConsole.MarkupLine(
         $"[yellow]The translation '{settings.Language.EscapeMarkup()}' does not define {missing.Count} " +
         $"{(missing.Count == 1 ? "label" : "labels")}, taken from '{TranslationsSelector.DefaultLanguage}': " +
         $"{string.Join(", ", missing).EscapeMarkup()}.[/]");
   }

   private static void WriteSanitizationReport(SanitizationReport report)
   {
      if (report.Corrections.Any())
      {
         AnsiConsole.MarkupLine(
            $"[yellow]The input file does not conform to the OpenAPI specification. Corrected {report.Corrections.Count} " +
            $"{(report.Corrections.Count == 1 ? "problem" : "problems")}:[/]");
         foreach (var correction in report.Corrections)
         {
            AnsiConsole.MarkupLine($"[yellow]  - {correction.ToString().EscapeMarkup()}[/]");
         }

         AnsiConsole.MarkupLine("[grey]  The corrections apply to the generated document only, the input file is not modified.[/]");
         AnsiConsole.MarkupLine("[grey]  Run with --no-sanitize to reject such a file instead.[/]");
      }

      foreach (var problem in report.SkippedProblems)
      {
         AnsiConsole.MarkupLine($"[red]Cannot be corrected - {problem.ToString().EscapeMarkup()}[/]");
      }
   }
}