using Spectre.Console.Cli;

namespace OpenApi2Excel.cli;

internal static class Program
{
   private static async Task<int> Main(string[] args)
   {
      var app = new CommandApp<GenerateExcelCommand>();

      app.Configure(config =>
      {
         config.SetHelpProvider(new CustomHelpProvider(config.Settings));
         config.SetApplicationName("OpenApi 2 Excel");
      });

      return await app.RunAsync(args);
   }
}