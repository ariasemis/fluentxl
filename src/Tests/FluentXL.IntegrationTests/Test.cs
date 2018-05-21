using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Validation;
using FluentXL.IntegrationTests.Utils;
using FluentXL.Specifications.Cells;
using FluentXL.Specifications.Columns;
using FluentXL.Specifications.MergeCells;
using FluentXL.Specifications.Rows;
using FluentXL.Specifications.Sheets;
using FluentXL.Writers;
using Microsoft.VisualStudio.TestTools.UnitTesting;
using System;
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
            var sheet = SheetSpecification
                .Sheet()
                .WithName("sheet 1")
                .WithColumn(
                    ColumnSpecification
                        .Column()
                        .With(index: 1, width: 100))
                .WithRow(
                    RowSpecification
                        .Row()
                        .WithIndex(1)
                        .WithCell(
                            CellSpecification
                                .Cell(1)
                                .WithColumn(1)
                                .WithContent("test")))
                .WithMergedCell(
                    MergeCellSpecification
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
                        SheetSpecification
                            .Sheet()
                            .WithName("test sheet")
                            .WithColumn(
                                ColumnSpecification
                                    .Column()
                                    .With(index: 1, width: 50))
                            .WithRow(
                                RowSpecification
                                    .Row()
                                    .WithIndex(2)
                                    .WithCell(
                                        CellSpecification
                                            .Cell(2)
                                            .WithColumn(2)
                                            .WithContent("value")))
                            .WithMergedCell(
                                MergeCellSpecification
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

        public void Method()
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
    }
}
