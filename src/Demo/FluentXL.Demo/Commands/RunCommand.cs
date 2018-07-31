using FluentXL.Demo.Utils;
using System;
using System.Collections.Generic;
using System.IO;
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
            Console.WriteLine("Generating sample");

            var sample = new Samples.Sample1();
            var filename = Path.Combine(FileHelper.GetOutputDirectory(), "sample.xlsx");

            using (var spreadsheet = new MemoryStream())
            using (var file = new FileStream(filename, FileMode.Create, FileAccess.Write))
            {
                //TODO: write on the filestream directly
                sample.Run().WriteTo(spreadsheet);

                spreadsheet.Seek(0, SeekOrigin.Begin);
                spreadsheet.CopyTo(file);
            }

            Console.WriteLine("sample generated");
            Console.WriteLine();

            return false;
        }
    }
}
