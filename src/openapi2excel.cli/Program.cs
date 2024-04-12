using openapi2excel.core;
using System.CommandLine;

namespace OpenApi2Excel.cli;

internal static class Program
{
   private static async Task<int> Main(string[] args)
   {
      var inputFileOption = new Option<FileInfo?>(
         name: "--file",
         description: "The path or URL to a YAML or JSON file with Rest API specification.",
         parseArgument: result =>
         {
            if (result.Tokens.Count == 0)
            {
               result.ErrorMessage = "Required argument missing for option: -f|--file.";
               return null;

            }
            var filePath = result.Tokens.Single().Value;
            if (File.Exists(filePath))
               return new FileInfo(filePath);
            if (Uri.IsWellFormedUriString(filePath, UriKind.RelativeOrAbsolute))
            {
               var inputFilePath = Path.Combine(Path.GetTempPath(), Path.GetRandomFileName() + ".xlsx");
               DownloadFileTaskAsync(new Uri(filePath), inputFilePath).GetAwaiter().GetResult();

               return new FileInfo(inputFilePath);
            }

            result.ErrorMessage = "File does not exist";
            return null;

         })
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
      var inputFilePath = inFile!.FullName;
      if (!inFile.Exists)
      {
         await Console.Error.WriteLineAsync("Invalid input file path.");
         return;
      }

      try
      {
         await OpenApiDocumentationGenerator
            .GenerateDocumentation(inputFilePath, outFile!.FullName, new OpenApiDocumentationOptions())
            .ConfigureAwait(false);

         Console.WriteLine("Excel file saved to " + outFile!.FullName);
      }
      catch (Exception exc)
      {
         await Console.Error.WriteLineAsync("An unexpected error occurred: " + exc.ToString());
      }
   }

   private static async Task DownloadFileTaskAsync(Uri uri, string fileName)
   {
      var client = new HttpClient();
      await using var s = await client.GetStreamAsync(uri);
      await using var fs = new FileStream(fileName, FileMode.CreateNew);
      await s.CopyToAsync(fs);
   }
}