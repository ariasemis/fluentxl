﻿using FluentXL.Demo.Samples;
using FluentXL.Demo.Utils;
using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Text;

namespace FluentXL.Demo.Commands
{
    public class RunCommand : ICommand
    {
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
                var filename = ResolveFilename();

                GenerateSample(sample, filename);
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
            //TODO
            return new Sample1();
        }

        private string ResolveFilename()
        {
            //TODO
            return Path.Combine(FileHelper.GetOutputDirectory(), "sample.xlsx");
        }
    }
}