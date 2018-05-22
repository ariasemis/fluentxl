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
    public class DocumentWriter
    {
        private IEnumerable<IBuilderSpecification<Sheet>> WorkSheets { get; set; }

        private DocumentWriter() { }

        public static DocumentWriter Create()
            => new DocumentWriter { WorkSheets = Enumerable.Empty<IBuilderSpecification<Sheet>>() };

        public DocumentWriter WithSheet(IBuilderSpecification<Sheet> specification)
        {
            return new DocumentWriter
            {
                WorkSheets = WorkSheets.Append(specification)
            };
        }

        public Stream Write()
        {
            var stream = new MemoryStream();

            using (var document = SpreadsheetDocument.Create(stream, SpreadsheetDocumentType.Workbook))
            {
                var workbookPart = document.AddWorkbookPart();

                var sheets = WriteWorksheet(workbookPart);
                WriteWorkbook(workbookPart, sheets);
            }

            stream.Seek(0, SeekOrigin.Begin);
            return stream;
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

        private List<OpenXml.Sheet> WriteWorksheet(WorkbookPart workbookPart)
        {
            var workbookSheets = new List<OpenXml.Sheet>();

            foreach (var sheet in WorkSheets.Select(x => x.Build()))
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
