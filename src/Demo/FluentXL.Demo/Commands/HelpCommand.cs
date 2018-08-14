using System;

namespace FluentXL.Demo.Commands
{
    public sealed class HelpCommand : ICommand
    {
        public bool Execute()
        {
            Console.WriteLine("Usage: [options] [command]");
            Console.WriteLine();

            Console.WriteLine("Options:");
            Console.WriteLine(" -h | -? | --help    Show help information");
            Console.WriteLine();

            Console.WriteLine("Commands:");
            Console.WriteLine(" run              executes the sample");
            Console.WriteLine(" sample           Show information about available samples");
            Console.WriteLine();

            Console.WriteLine("Use '[command] --help' for more information about a command.");

            return false;
        }
    }
}
