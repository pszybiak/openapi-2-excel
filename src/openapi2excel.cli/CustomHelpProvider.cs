using Spectre.Console;
using Spectre.Console.Cli;
using Spectre.Console.Cli.Help;
using Spectre.Console.Rendering;

namespace OpenApi2Excel.cli;

internal class CustomHelpProvider(ICommandAppSettings settings) : HelpProvider(settings)
{
   public override IEnumerable<IRenderable> GetHeader(ICommandModel model, ICommandInfo? command)
      => GetHeaderText();

   public static IEnumerable<IRenderable> GetHeaderText()
      => new[]
      {
         Text.NewLine, Text.NewLine,
         new Text("    ██████╗ ██████╗ ███████╗███╗   ██╗ █████╗ ██████╗ ██╗    ██████╗     ███████╗██╗  ██╗ ██████╗███████╗██╗"), Text.NewLine,
         new Text("   ██╔═══██╗██╔══██╗██╔════╝████╗  ██║██╔══██╗██╔══██╗██║    ╚════██╗    ██╔════╝╚██╗██╔╝██╔════╝██╔════╝██║"), Text.NewLine,
         new Text("   ██║   ██║██████╔╝█████╗  ██╔██╗ ██║███████║██████╔╝██║     █████╔╝    █████╗   ╚███╔╝ ██║     █████╗  ██║"), Text.NewLine,
         new Text("   ██║   ██║██╔═══╝ ██╔══╝  ██║╚██╗██║██╔══██║██╔═══╝ ██║    ██╔═══╝     ██╔══╝   ██╔██╗ ██║     ██╔══╝  ██║"), Text.NewLine,
         new Text("   ╚██████╔╝██║     ███████╗██║ ╚████║██║  ██║██║     ██║    ███████╗    ███████╗██╔╝ ██╗╚██████╗███████╗███████╗"), Text.NewLine,
         new Text("    ╚═════╝ ╚═╝     ╚══════╝╚═╝  ╚═══╝╚═╝  ╚═╝╚═╝     ╚═╝    ╚══════╝    ╚══════╝╚═╝  ╚═╝ ╚═════╝╚══════╝╚══════╝"), Text.NewLine,
         Text.NewLine,Text.NewLine,
      };
}