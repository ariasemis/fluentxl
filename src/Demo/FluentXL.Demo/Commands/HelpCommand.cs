using System;

namespace FluentXL.Demo.Commands
{
    public sealed class HelpCommand : ICommand
    {
        public bool Execute()
        {
            Console.WriteLine("Usage: run [arguments] [options]");
            Console.WriteLine();
            Console.WriteLine("Arguments:");
            Console.WriteLine(" sample              String         name of the sample you want to run");
            Console.WriteLine();
            Console.WriteLine("Options:");
            Console.WriteLine(" -h | -? | --help    Show help information");
            Console.WriteLine(" -f | --file         String         File path where the sample will be generated");
            Console.WriteLine();

            return false;
        }
    }
}
