using FluentXL.Demo.Commands;
using System;
using System.Collections.Generic;

namespace FluentXL.Demo
{
    public static class Parser
    {
        public static ICommand Parse(string raw)
        {
            if (string.IsNullOrWhiteSpace(raw))
                throw new ArgumentException("raw input must have a value", nameof(raw));

            var input = GetInput(raw);

            switch (input.Command.ToLower())
            {
                //TODO: handle common exit codes
                case "exit":
                case "quit":
                case "close":
                    return new ExitCommand();
                case "help":
                case "--help":
                case "-h":
                case "-?":
                    return new HelpCommand();
                case "run":
                    return new RunCommand(input);
                default:
                    return new UnknownCommand(input.Command);
            }
        }

        private static Input GetInput(string input)
        {
            //TODO: add more robust implementation
            var parts = input.Trim().Split(' ');

            var result = new Input
            {
                Command = parts[0]
            };

            if (parts.Length > 1)
            {
                var i = 1;
                while (i < parts.Length)
                {
                    var part = parts[i];
                    if (IsOption(part))
                    {
                        // check if the option is followed by an argument
                        var j = i + 1;
                        var next = j < parts.Length ? parts[j] : null;

                        if (string.IsNullOrEmpty(next) || IsOption(next))
                        {
                            // if not, is a flag
                            result.Flags.Add(part);
                        }
                        else
                        {
                            // we are not supporting options with multiple arguments
                            result.Options.Add(part, next);
                            i = j;
                        }
                    }
                    else
                    {
                        result.Arguments.Add(part);
                    }

                    i++;
                }
            }

            return result;
        }

        private static bool IsOption(string input)
            => input.StartsWith("-");
    }
}
