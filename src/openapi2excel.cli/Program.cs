using openapi2excel.core;
using System.CommandLine;

namespace OpenApi2Excel.cli;

internal static class Program
{
   private static async Task<int> Main(string[] args)
   {
      var inputFileOption = new Option<FileInfo?>(
         name: "--file",
         description: "The path to a YAML or JSON file with Rest API specification.")
      { IsRequired = true };
      inputFileOption.AddAlias("-f");

      var outputFileOption = new Option<FileInfo?>(
         name: "--out",
         description: "The path for output excel file.")
      { IsRequired = true };
      outputFileOption.AddAlias("-o");

      var rootCommand = new RootCommand("OpenApi-2-Excel");
      rootCommand.AddOption(inputFileOption);
      rootCommand.AddOption(outputFileOption);
      rootCommand.SetHandler(HandleFileGeneration, inputFileOption, outputFileOption);

      return await rootCommand.InvokeAsync(args);
   }

   private static async Task HandleFileGeneration(FileInfo? inFile, FileInfo? outFile)
   {
      if (!inFile!.Exists)
      {
         await Console.Error.WriteLineAsync("Invalid input file path.");
         return;
      }

      try
      {
         await OpenApiDocumentationGenerator
            .GenerateDocumentation(inFile!.FullName, outFile!.FullName, new OpenApiDocumentationOptions())
            .ConfigureAwait(false);
      }
      catch (Exception exc)
      {
         await Console.Error.WriteLineAsync("An unexpected error occurred: " + exc.Message);
      }
   }
}