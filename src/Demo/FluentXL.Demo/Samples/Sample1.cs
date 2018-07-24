namespace FluentXL.Demo.Samples
{
    /// <summary>
    /// This sample shows how to create a new workbook from scratch.
    /// </summary>
    public class Sample1 : ISample
    {
        public ISpreadsheetWriter Run()
        {
            // define header
            var headerStyle = Specification.CellFormat()
                .WithFill(Specification.Fill()
                    .WithPattern(Models.FillPattern.Solid)
                    .WithBackgroundColor(Specification.Color().FromArgb(0, 51, 153, 255)))
                .WithFont(Specification.Font()
                    .WithFont("Arial")
                    .WithSize(10)
                    .Bold()
                    .WithColor(Specification.Color().FromArgb("FFFFFFFF")));

            var header = Specification.Row()
                .OnIndex(1)
                .WithCell(Specification.Cell().OnColumn(1).WithContent("CLIENT INFORMATION").WithStyle(headerStyle))
                .WithCell(Specification.Cell().OnColumn(3).WithContent("OTHER INFORMATION").WithStyle(headerStyle));

            // define sub header
            var subHeader = Specification.Row()
                .OnIndex(2)
                .WithCell(Specification.Cell().OnColumn(1).WithContent("Code").WithStyle(headerStyle))
                .WithCell(Specification.Cell().OnColumn(2).WithContent("Name").WithStyle(headerStyle))
                .WithCell(Specification.Cell().OnColumn(3).WithContent("Debt").WithStyle(headerStyle));

            // define data rows
            // TODO:

            // define document
            var doc = Spreadsheet.New()
                .WithSheet(Specification.Sheet()
                    .WithName("sample 1")
                    .WithRow(header)
                    .WithRow(subHeader));

            return doc;
        }
    }
}
