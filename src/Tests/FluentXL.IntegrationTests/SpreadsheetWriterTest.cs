using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Validation;
using FluentXL.IntegrationTests.Utils;
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
                                            .WithContent("Hello World!!")
                                            .WithStyle(
                                                Specification
                                                    .CellFormat()
                                                    .WithFill(Specification
                                                        .Fill()
                                                        .WithPattern(Models.FillPattern.Solid)
                                                        .WithBackgroundColor(Specification.Color().FromRgb("FF696969"))
                                                        .WithForegroundColor(Specification.Color().FromRgb("FF696969")))
                                                    .WithBorder(Specification
                                                        .Border()
                                                        .WithOutline(
                                                            Models.BorderStyle.Medium,
                                                            Specification.Color().FromRgb("#00000000"))))))
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

    }
}
