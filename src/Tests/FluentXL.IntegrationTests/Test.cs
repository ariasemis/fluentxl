using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Validation;
using FluentXL.IntegrationTests.Utils;
using FluentXL.Specifications;
using FluentXL.Writers;
using Microsoft.VisualStudio.TestTools.UnitTesting;
using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;

namespace FluentXL.IntegrationTests
{
    [TestClass]
    public class Test
    {
        [TestInitialize]
        public void Init()
        {
            if (!Directory.Exists(FileHelper.GetOutputDirectory()))
                Directory.CreateDirectory(FileHelper.GetOutputDirectory());
        }

        [TestMethod]
        public void Other()
        {
            var sheet = Specification
                .Sheet()
                .WithName("sheet 1")
                .WithColumn(
                    Specification
                        .Column()
                        .With(index: 1, width: 100))
                .WithRow(
                    Specification
                        .Row()
                        .OnIndex(1)
                        .WithCell(
                            Specification
                                .Cell()
                                .OnColumn(1)
                                .WithContent("test")))
                .WithMergedCell(
                    Specification
                        .MergeCell()
                        .From(row: 1, column: 1)
                        .To(row: 1, column: 2))
                .Build();

            Assert.IsNotNull(sheet);
            Assert.AreEqual("sheet 1", sheet.Name);
            Assert.IsNotNull(sheet.Columns);
            Assert.IsNotNull(sheet.Columns?.SingleOrDefault());
            Assert.IsNotNull(sheet.Rows);
            Assert.IsNotNull(sheet.Rows?.SingleOrDefault());
            Assert.IsNotNull(sheet.Rows?.SingleOrDefault()?.Cells);
            Assert.IsNotNull(sheet.MergeCells);
            Assert.IsNotNull(sheet.MergeCells?.SingleOrDefault());
            Assert.AreEqual(1, sheet.MergeCells?.Count);
        }

        [TestMethod]
        public void CreateExcel()
        {
            // arrange
            var filename = Path.Combine(FileHelper.GetOutputDirectory(), "test.xlsx");

            if (FileHelper.IsFileLocked(filename))
            {
                Assert.Inconclusive("test cannot be performed because the output file is locked");
            }

            Func<Stream> export =
                DocumentWriter
                    .Create()
                    .WithSheet(
                        Specification
                            .Sheet()
                            .WithName("test sheet")
                            .WithColumn(
                                Specification
                                    .Column()
                                    .With(index: 1, width: 50))
                            .WithRow(
                                Specification
                                    .Row()
                                    .OnIndex(2)
                                    .WithCell(
                                        Specification
                                            .Cell()
                                            .OnColumn(2)
                                            .WithContent("Hello World!!")))
                            .WithMergedCell(
                                Specification
                                    .MergeCell()
                                    .From(row: 2, column: 2)
                                    .To(row: 2, column: 3)))
                    .Write;

            // act
            using (var spreadsheet = export())
            using (var file = new FileStream(filename, FileMode.Create, FileAccess.Write))
            {
                spreadsheet.Seek(0, SeekOrigin.Begin);
                spreadsheet.CopyTo(file);
            }

            // assert
            using (var document = SpreadsheetDocument.Open(filename, false))
            {
                var validator = new OpenXmlValidator();
                var errors = validator.Validate(document);

                Assert.IsNotNull(errors);
                Assert.IsFalse(errors.Any());
            }
        }

        [TestMethod]
        public void CreateBigExcel()
        {
            // arrange
            var filename = Path.Combine(FileHelper.GetOutputDirectory(), "big.xlsx");

            if (FileHelper.IsFileLocked(filename))
            {
                Assert.Inconclusive("test cannot be performed because the output file is locked");
            }

            var data = GetData(1000);
            var document = DocumentWriter
                .Create()
                .WithSheet(
                    Specification
                        .Sheet()
                        .WithName("sheet 1")
                        .WithColumns(Enumerable.Range(0, 5), (item, index) => Specification.Column().With(index: index, width: 30))
                        .WithRows(data, (item, index) => Specification
                            .Row()
                            .OnIndex(index)
                            .WithCell(Specification.Cell().OnColumn(1).WithContent(item.Name))
                            .WithCell(Specification.Cell().OnColumn(2).WithContent(item.Description))
                            .WithCell(Specification.Cell().OnColumn(3).WithContent(item.Number))
                            .WithCell(Specification.Cell().OnColumn(4).WithContent(item.Money))
                            .WithCell(Specification.Cell().OnColumn(5).WithContent(item.Date))
                            ));

            // act
            using (var spreadsheet = document.Write())
            using (var file = new FileStream(filename, FileMode.Create, FileAccess.Write))
            {
                spreadsheet.Seek(0, SeekOrigin.Begin);
                spreadsheet.CopyTo(file);
            }

            // assert
            using (var doc = SpreadsheetDocument.Open(filename, false))
            {
                var validator = new OpenXmlValidator();
                var errors = validator.Validate(doc);

                Assert.IsNotNull(errors);
                Assert.IsFalse(errors.Any());
            }
        }

        private IEnumerable<Dummy> GetData(int size)
        {
            for (int i = 0; i < size; i++)
                yield return new Dummy
                {
                    Name = "something",
                    Description = "Lorem ipsum dolor sit amet, consectetur adipiscing elit. Proin id maximus justo, tempor feugiat quam. Praesent turpis nunc, semper non neque nec, semper scelerisque nulla.",
                    Number = 999,
                    Money = 99.9M,
                    Date = DateTime.Today
                };
        }

        private void Contract()
        {
            //TODO:

            //desired client interface
            dynamic builder = null;
            dynamic sheetSpecification = null;
            dynamic columnSpecification = null;
            dynamic rowSpecification = null;
            dynamic cellSpecification = null;
            dynamic cellStyleSpecification = null;
            dynamic mergeCellSpecification = null;

            // usage
            dynamic result =
                builder
                    .WithSheet(
                        sheetSpecification
                            .WithName("sheet 1")
                            .WithColumn(
                                columnSpecification.With(index: 1, width: 100))
                            //.WithColumn(
                            //    () => new { index = 1, width = 100 })
                            //.WithColumns(
                            //    source, c => columnSpecification.With(c.index, c.width));
                            .WithRow(
                                rowSpecification
                                    .WithIndex(1)
                                    .WithCell(
                                        cellSpecification
                                            .WithColumn(1)
                                            .WithContent("value")
                                            .WithStyle(
                                                cellStyleSpecification
                                                    .WithFont()
                                                    .WithBackgroundColor()
                                                    .WithForegroundColor()
                                                    .WithBorder()
                                                    .WithNumberFormat()
                                                    .WithAlignment()))
                                    .WithCells())
                            .WithRows()
                            .WithMergedCell(
                                mergeCellSpecification
                                    .From(row: 1, column: 1)
                                    .To(row: 2, column: 1))
                            .WithMergedCells())
                    .Build();

            Assert.Fail();
        }

        class Dummy
        {
            public string Name { get; set; }
            public string Description { get; set; }
            public int Number { get; set; }
            public decimal Money { get; set; }
            public DateTime Date { get; set; }
        }
    }
}
