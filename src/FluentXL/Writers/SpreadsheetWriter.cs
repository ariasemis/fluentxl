using DocumentFormat.OpenXml;
using DocumentFormat.OpenXml.Packaging;
using FluentXL.Models;
using FluentXL.Specifications;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using OpenXml = DocumentFormat.OpenXml.Spreadsheet;

namespace FluentXL.Writers
{
    public class SpreadsheetWriter : ISpreadsheetWriter
    {
        private IEnumerable<IBuilderSpecification<Sheet>> WorkSheets { get; set; }

        public SpreadsheetWriter()
        {
            WorkSheets = Enumerable.Empty<IBuilderSpecification<Sheet>>();
        }

        public ISpreadsheetWriter WithSheet(IBuilderSpecification<Sheet> specification)
        {
            return new SpreadsheetWriter
            {
                WorkSheets = WorkSheets.Append(specification)
            };
        }

        public void WriteTo(Stream stream)
        {
            using (var document = SpreadsheetDocument.Create(stream, SpreadsheetDocumentType.Workbook))
            {
                var context = new BuildContext();
                var workbookPart = document.AddWorkbookPart();

                var sheets = WriteWorksheet(workbookPart, context);
                WriteWorkbookStyles(workbookPart, context);
                WriteWorkbook(workbookPart, sheets);
            }
        }

        private static void WriteWorkbook(WorkbookPart workbookPart, List<OpenXml.Sheet> sheets)
        {
            using (var workbookWriter = OpenXmlWriter.Create(workbookPart))
            {
                workbookWriter.WriteStartElement(new OpenXml.Workbook());
                workbookWriter.WriteStartElement(new OpenXml.Sheets());
                foreach (var sheet in sheets)
                {
                    workbookWriter.WriteElement(sheet);
                }
                workbookWriter.WriteEndElement();
                workbookWriter.WriteEndElement();
                workbookWriter.Close();
            }
        }

        private static void WriteWorkbookStyles(WorkbookPart workbookPart, IBuildContext context)
        {
            var stylesPart = workbookPart.AddNewPart<WorkbookStylesPart>();
            var stylesheetWriter = new StylesheetWriter(stylesPart);

            var stylesheet = context.Stylesheet.Build();

            stylesheetWriter.Write(stylesheet);
        }

        private List<OpenXml.Sheet> WriteWorksheet(WorkbookPart workbookPart, IBuildContext context)
        {
            var workbookSheets = new List<OpenXml.Sheet>();

            foreach (var sheet in WorkSheets.Select(x => x.Build(context)))
            {
                var worksheetPart = workbookPart.AddNewPart<WorksheetPart>();

                workbookSheets.Add(new OpenXml.Sheet
                {
                    SheetId = (uint)workbookSheets.Count + 1,
                    Name = sheet.Name,
                    Id = workbookPart.GetIdOfPart(worksheetPart)
                });

                var sheetWriter = new SheetWriter(worksheetPart);
                sheetWriter.Write(sheet);
            }

            return workbookSheets;
        }
    }
}
