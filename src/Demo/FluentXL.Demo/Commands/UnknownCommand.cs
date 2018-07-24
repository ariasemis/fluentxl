using System;

namespace FluentXL.Demo.Commands
{
    public class UnknownCommand : ICommand
    {
        private readonly string name;

        public UnknownCommand(string name)
        {
            this.name = name;
        }

        public bool Execute()
        {
            Console.WriteLine($"'{name}' is not valid command. See 'help'.");
            Console.WriteLine();

            return false;
        }
    }
}
