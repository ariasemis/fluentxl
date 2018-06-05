using BenchmarkDotNet.Attributes;
using FluentXL.Specifications;
using FluentXL.Writers;
using System;
using System.Collections.Generic;
using System.IO;

namespace FluentXL.Benchmarks
{
    public class SpreadsheetWriterBenchmark
    {
        private SpreadsheetWriter documentWriter;

        [Params(100, 1000, 10000)]
        public int RowSize { get; set; }

        [GlobalSetup]
        public void Setup()
        {
            var data = GetData(RowSize);

            documentWriter = Spreadsheet
                .New()
                .WithSheet(
                    Specification.Sheet()
                        .WithName("benchmark")
                        .WithRows(data, (item, index) => Specification
                            .Row()
                            .OnIndex(index)
                            .WithCell(Specification.Cell().OnColumn(1).WithContent(item.Text))
                            .WithCell(Specification.Cell().OnColumn(2).WithContent(item.Number))
                            .WithCell(Specification.Cell().OnColumn(3).WithContent(item.Money))
                            .WithCell(Specification.Cell().OnColumn(4).WithContent(item.Date))
                            ));
        }

        [Benchmark]
        public void Measure_DocumentWriter_Write()
        {
            using (var stream = new MemoryStream())
            {
                documentWriter.WriteTo(stream);
            }
        }

        private IEnumerable<Data> GetData(int size)
        {
            for (int i = 0; i < size; i++)
                yield return new Data
                {
                    Text = "something",
                    Number = 999,
                    Money = 99.9M,
                    Date = DateTime.Today
                };
        }

        class Data
        {
            public string Text { get; set; }
            public int Number { get; set; }
            public decimal Money { get; set; }
            public DateTime Date { get; set; }
        }
    }
}
