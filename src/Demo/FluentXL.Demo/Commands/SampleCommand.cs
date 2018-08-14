using FluentXL.Demo.Extensions;
using FluentXL.Demo.Samples;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Reflection;

namespace FluentXL.Demo.Commands
{
    public class SampleCommand : ICommand
    {
        private readonly Input input;

        public SampleCommand(Input input)
        {
            this.input = input ?? throw new ArgumentNullException(nameof(input));
        }

        public bool Execute()
        {
            var sampleInput = input.Options
                .Where(x => IsName(x.Key))
                .Select(x => (KeyValuePair<string, string>?)x)
                .FirstOrDefault();

            if (sampleInput.HasValue && sampleInput.Value.Value != null)
            {
                PrintSampleInfo(sampleInput.Value.Value);
            }
            else
            {
                if (input.Flags.Any(x => IsHelp(x)))
                {
                    PrintHelp();
                }
                else
                {
                    PrintSamples();
                }
            }

            return false;
        }

        private static bool IsHelp(string input)
            => new[] { "-h", "-?", "--help" }.Contains(input);

        private static bool IsName(string input)
            => new[] { "-n", "--name" }.Contains(input);

        private void PrintSampleInfo(string name)
        {
            ISample sample = Assembly.GetExecutingAssembly().ResolveSample(name);

            if (sample == null)
            {
                Console.WriteLine("sample name not recognized");
                Console.WriteLine("Run 'sample' command for a list of available samples");
                Console.WriteLine();
            }
            else
            {
                Console.WriteLine(sample.GetInfo());
                Console.WriteLine();
            }
        }

        private void PrintHelp()
        {
            Console.WriteLine("Usage: sample [arguments] [options]");
            Console.WriteLine();
            Console.WriteLine("Options:");
            Console.WriteLine(" -h | -? | --help    Show help information");
            Console.WriteLine(" -n | --name         String         Name of the sample you want to know more");
            Console.WriteLine();
        }

        private void PrintSamples()
        {
            var samples = Assembly
                .GetExecutingAssembly()
                .GetSampleNames();

            Console.WriteLine("Available samples:");

            foreach (var sample in samples)
            {
                Console.WriteLine($" - {sample}");
            }

            Console.WriteLine();
        }
    }
}
