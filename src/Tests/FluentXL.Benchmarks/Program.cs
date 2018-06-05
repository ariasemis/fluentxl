using BenchmarkDotNet.Running;
using System;

namespace FluentXL.Benchmarks
{
    class Program
    {
        static void Main(string[] args)
        {
            var summary = BenchmarkRunner.Run<SpreadsheetWriterBenchmark>();

            Console.ReadLine();
        }
    }
}
