using System;
using System.Collections.Generic;
using System.CommandLine;
using System.CommandLine.Invocation;
using System.IO;
using System.Linq;
using System.Threading.Tasks;

namespace SteffBeckers.Abp.Cli.Localization
{
    public class Program
    {
        public static async Task<int> Main(string[] args)
        {
            RootCommand rootCommand = new RootCommand("Steff's ABP.io Localization CLI");

            Command listCommand = new Command("list", "Lists all localization files in directory.");
            listCommand.Handler = CommandHandler.Create(() =>
            {
                List<string> filePaths = Directory.GetFiles(Directory.GetCurrentDirectory())
                    .Where(x => x.EndsWith(".json"))
                    .ToList();

                filePaths.ForEach(x => Console.WriteLine(x));
            });
            rootCommand.AddCommand(listCommand);

            return await rootCommand.InvokeAsync(args);
        }
    }
}