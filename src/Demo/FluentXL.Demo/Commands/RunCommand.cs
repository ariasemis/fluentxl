using FluentXL.Demo.Extensions;
using FluentXL.Demo.Samples;
using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Reflection;

namespace FluentXL.Demo.Commands
{
    public class RunCommand : ICommand
    {
        private const string DEFAULT_FILE_NAME = "sample.xlsx";
        private const string FILE_EXTENSION = ".xlsx";

        private readonly Input input;

        public RunCommand(Input input)
        {
            this.input = input ?? throw new ArgumentNullException(nameof(input));
        }

        public bool Execute()
        {
            if (input.Flags.Any(x => IsHelp(x)))
            {
                ShowHelp();
            }
            else
            {
                var sample = ResolveSample();

                if (sample != null)
                {
                    var filename = ResolveFilename();

                    GenerateSample(sample, filename);
                }
            }

            return false;
        }

        private static bool IsHelp(string input)
            => new[] { "-h", "-?", "--help" }.Contains(input);

        private static bool IsFile(string input)
            => new[] { "-f", "--file" }.Contains(input);

        private void ShowHelp()
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
        }

        private void GenerateSample(ISample sample, string filename)
        {
            Console.WriteLine("Generating sample...");

            //TODO: write on the filestream directly instead of using memory stream
            using (var spreadsheet = new MemoryStream())
            using (var file = new FileStream(filename, FileMode.Create, FileAccess.Write))
            {
                sample.Run().WriteTo(spreadsheet);

                spreadsheet.Seek(0, SeekOrigin.Begin);
                spreadsheet.CopyTo(file);
            }

            Console.WriteLine($"Sample generated at '{filename}'");
            Console.WriteLine();
        }

        private ISample ResolveSample()
        {
            ISample sample;

            if (input.Arguments.Any())
            {
                var inputName = input.Arguments.First();

                sample = Assembly
                    .GetExecutingAssembly()
                    .ResolveSample(inputName);

                if (sample == null)
                {
                    Console.WriteLine($"sample with name = '{inputName}' not found");
                    Console.WriteLine("sample generation was cancelled");
                    Console.WriteLine();
                    Console.WriteLine("See 'sample' command for a list of available samples");
                }
            }
            else
            {
                sample = new Sample1();
            }

            return sample;
        }

        private string ResolveFilename()
        {
            var result = string.Empty;

            var fileInput = input.Options
                .Where(x => IsFile(x.Key))
                .Select(x => (KeyValuePair<string, string>?)x)
                .FirstOrDefault();

            if (!fileInput.HasValue || fileInput.Value.Value == null)
            {
                result = Path.GetFullPath(DEFAULT_FILE_NAME);
            }
            else
            {
                var path = fileInput.Value.Value.Trim();

                var filename = Path.GetFileName(path);
                var directory = Path.GetDirectoryName(path);

                if (string.IsNullOrWhiteSpace(filename))
                {
                    filename = DEFAULT_FILE_NAME;
                }
                else
                {
                    if (Path.HasExtension(filename))
                    {
                        if (Path.GetExtension(filename) != FILE_EXTENSION)
                            throw new ArgumentException("file extension not supported");
                    }
                    else
                    {
                        filename += ".xlsx";
                    }
                }

                result = Path.Combine(directory, filename);
            }

            return result;
        }
    }
}
