using DocumentFormat.OpenXml;
using DocumentFormat.OpenXml.Packaging;
using FluentXL.Models;
using System;
using System.Collections.Generic;
using System.IO;
using System.Text;
using OpenXml = DocumentFormat.OpenXml.Spreadsheet;

namespace FluentXL.Writers
{
    internal class DocumentWriter
    {
        IEnumerable<Sheet> worksheets = null;

        public Stream Write()
        {
            var stream = new MemoryStream();
            var document = SpreadsheetDocument.Create(stream, SpreadsheetDocumentType.Workbook);
            var workbookPart = document.AddWorkbookPart();

            // WRITE WORKSHEETS

            // keep track of worksheets to reference later in workbook
            var workbookSheets = new List<OpenXml.Sheet>();

            foreach (var sheet in worksheets)
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

            // WRITE WORKBOOK
            var workbookWriter = OpenXmlWriter.Create(workbookPart);
            workbookWriter.WriteStartElement(new OpenXml.Workbook());
            workbookWriter.WriteStartElement(new OpenXml.Sheets());
            foreach (var sheet in workbookSheets)
            {
                workbookWriter.WriteElement(sheet);
            }
            workbookWriter.WriteEndElement();
            workbookWriter.WriteEndElement();
            workbookWriter.Close();

            // END
            document.Dispose();
            stream.Seek(0, SeekOrigin.Begin);
            return stream;
        }
    }
}
