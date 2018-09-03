using FluentXL.Models;
using System;

namespace FluentXL.Demo.Samples
{
    /// <summary>
    /// This sample shows how to create a new workbook from scratch.
    /// </summary>
    public class Showcase : ISample
    {
        public string GetInfo()
        {
            return "Creates a basic spreadsheet with a header, a subheader and a list of data";
        }

        public ISpreadsheetWriter Run()
        {
            // we start by defining the style for the header rows
            var headerStyle = Specification.CellFormat()
                .WithFont(Specification.Font()
                    .WithFont("Arial")
                    .WithSize(10)
                    .Bold())
                .WithBorder(Specification.Border().WithOutline(
                    BorderStyle.Thin,
                    Specification.Color().FromArgb("FF000000")))
                .WithFill(Specification.Fill()
                    .WithPattern(FillPattern.Solid)
                    .WithForegroundColor(Specification.Color().FromArgb("FFC6EFCE")));

            // we'll use the first couple of rows as a kind of "header" where we'll display some titles for the columns
            // for the first row, we group the columns in 2 groups depending on the kind of data displayed
            // so we create a title for each group and add them to the first row
            var header = Specification.Row()
                .OnIndex(1)
                .WithCell(Specification.Cell().OnColumn(1).WithContent("CLIENT INFORMATION").WithStyle(headerStyle))
                .WithCell(Specification.Cell().OnColumn(4).WithContent("OTHER INFORMATION").WithStyle(headerStyle));

            // in the second row we display a title for each column
            var subHeader = Specification.Row()
                .OnIndex(2)
                .WithCell(Specification.Cell().OnColumn(1).WithContent("Code").WithStyle(headerStyle))
                .WithCell(Specification.Cell().OnColumn(2).WithContent("Name").WithStyle(headerStyle))
                .WithCell(Specification.Cell().OnColumn(3).WithContent("Date").WithStyle(headerStyle))
                .WithCell(Specification.Cell().OnColumn(4).WithContent("Debt").WithStyle(headerStyle));

            // now, we start by defining the result data we want to export in the excel file
            // in this case, we just hardcoded an array of objects
            var data = new[]
            {
                new { Code = "#0000001", Date = new DateTime(1990, 1, 1), Name = "John Smith", Debt = 1000M },
                new { Code = "#0000240", Date = new DateTime(1990, 6, 13), Name = "John Doe", Debt = 570.33M },
                new { Code = "#0000367", Date = new DateTime(1990, 10, 7), Name = "Jane Doe", Debt = 1706M },
                new { Code = "#0000398", Date = new DateTime(1990, 2, 14), Name = "Advent corporation", Debt = 37080.16M },
                new { Code = "#0001002", Date = new DateTime(1990, 11, 7), Name = "Morley", Debt = 70180.20M },
                new { Code = "#0001031", Date = new DateTime(1990, 12, 31), Name = "Newco", Debt = 1040M },
                new { Code = "#0001108", Date = new DateTime(1990, 3, 31), Name = "XYZ Widget Company", Debt = 64870.21M },
                new { Code = "#0001578", Date = new DateTime(1990, 7, 15), Name = "Contoso", Debt = 43657M },
                new { Code = "#0002096", Date = new DateTime(1990, 9, 10), Name = "Foo Bar", Debt = 890M },
                new { Code = "#0002201", Date = new DateTime(1990, 11, 18), Name = "J. Random X", Debt = 11098.75M }
            };

            // finally we tie everything together:
            // we create a document with one sheet
            // add the previously created rows
            // then create the rows for the data exported
            // and at the end we merge the cells corresponding to the header groups
            var doc = Spreadsheet.New()
                .WithSheet(Specification.Sheet()
                    .WithName("sample 1")
                    .WithRow(header)
                    .WithRow(subHeader)
                    .WithRows(data, (d, i) => Specification.Row()
                        .OnIndex(i + 2)
                        .WithCell(Specification.Cell().OnColumn(1).WithContent(d.Code))
                        .WithCell(Specification.Cell().OnColumn(2).WithContent(d.Name))
                        .WithCell(Specification.Cell().OnColumn(3).WithContent(d.Date))
                        .WithCell(Specification.Cell().OnColumn(4).WithContent(d.Debt)))
                    .WithColumns(new[]
                    {
                        Specification.Column().With(index: 1, width: 12),
                        Specification.Column().With(index: 2, width: 18),
                        Specification.Column().With(index: 3, width: 15),
                        Specification.Column().With(index: 4, width: 23)
                    })
                    .WithMergedCell(Specification.MergeCell().From(1, 1).To(1, 3)));

            // the exported excel should look something like this:
            /* -----------------------------------------------------
             * | CLIENT INFORMATION            | OTHER INFORMATION |
             * -----------------------------------------------------
             * | Code   | Name   | Date        | Debt              |
             * -----------------------------------------------------
             * | 001    | Smith  | 1/1/1990    | 1000              |
             * -----------------------------------------------------
             * | ...    | ...    | ...         | ...               |
             * -----------------------------------------------------
             * */
            return doc;
        }
    }
}
