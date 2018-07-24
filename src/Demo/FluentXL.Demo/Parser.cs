using FluentXL.Demo.Commands;
using System;
using System.Linq;

namespace FluentXL.Demo
{
    public static class Parser
    {
        public static ICommand Parse(string input)
        {
            if (string.IsNullOrWhiteSpace(input))
                throw new ArgumentException("input must have a value", nameof(input));

            var parts = GetInputParts(input);

            switch (parts.Command.ToLower())
            {
                //TODO: handle common exit codes
                case "exit":
                    return new ExitCommand();
                case "help":
                case "--help":
                case "-h":
                case "-?":
                    return new HelpCommand();
                case "run":
                    return new RunCommand();
                default:
                    return new UnknownCommand(parts.Command);
            }
        }

        private static CommandParts GetInputParts(string input)
        {
            //TODO: add more robust implementation
            var parts = input.Split(' ');

            return new CommandParts
            {
                Command = parts[0],
                Arguments = parts.Skip(1).ToArray()
            };
        }

        class CommandParts
        {
            public string Command { get; set; }
            public string[] Arguments { get; set; }
        }
    }
}
