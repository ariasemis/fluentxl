using System;

namespace FluentXL.Demo
{
    class Program
    {
        private static readonly object syncLock = new object();
        private static bool exitRequested = false;

        static void Main(string[] args)
        {
            // set up
            Console.CancelKeyPress += Console_CancelKeyPress;

            // display info
            Console.WriteLine("Welcome to the FluentXL demo application.");
            Console.WriteLine("See 'help' for more information.");
            Console.WriteLine();

            // run
            while (!exitRequested)
            {
                Console.Write("> ");

                var input = Console.ReadLine();
                if (string.IsNullOrWhiteSpace(input))
                    continue;

                var command = Parser.Parse(input);
                var exit = command.Execute();

                if (exit) RequestExit();
            }
        }

        private static void Console_CancelKeyPress(object sender, ConsoleCancelEventArgs e)
        {
            RequestExit();
            Environment.Exit(-1);
        }

        private static void RequestExit()
        {
            lock (syncLock)
            {
                exitRequested = true;
            }
        }
    }
}
