using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Validation;
using FluentXL.IntegrationTests.Utils;
using FluentXL.Specifications;
using Microsoft.VisualStudio.TestTools.UnitTesting;
using System.IO;
using System.Linq;

namespace FluentXL.IntegrationTests
{
    [TestClass]
    public class SpreadsheetWriterTest
    {
        [TestInitialize]
        public void Init()
        {
            if (!Directory.Exists(FileHelper.GetOutputDirectory()))
                Directory.CreateDirectory(FileHelper.GetOutputDirectory());
        }

        [TestMethod]
        public void Write_WithValidData_Succeeds()
        {
            // arrange
            var filename = Path.Combine(FileHelper.GetOutputDirectory(), "test.xlsx");

            if (FileHelper.IsFileLocked(filename))
            {
                Assert.Inconclusive("test cannot be performed because the output file is locked");
            }

            var doc =
                Spreadsheet
                    .New()
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
                                    .To(row: 2, column: 3)));

            // act
            using (var spreadsheet = new MemoryStream())
            using (var file = new FileStream(filename, FileMode.Create, FileAccess.Write))
            {
                //TODO: write on the filestream directly
                doc.WriteTo(spreadsheet);

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
    }
}
