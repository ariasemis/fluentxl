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

                WriteWorkbookStyles(workbookPart);
                var sheets = WriteWorksheet(workbookPart, context);
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

        private static void WriteWorkbookStyles(WorkbookPart workbookPart)
        {
            var stylesheet = CreateStylesheet();

            var stylesPart = workbookPart.AddNewPart<WorkbookStylesPart>();
            stylesPart.Stylesheet = stylesheet;

            stylesheet.Save();
        }

        private static OpenXml.Stylesheet CreateStylesheet()
        {
            var stylesheet = new OpenXml.Stylesheet();
            stylesheet.Append(new OpenXml.Fonts { Count = 0 });
            stylesheet.Append(new OpenXml.Fills { Count = 0 });
            stylesheet.Append(new OpenXml.Borders { Count = 0 });
            //stylesheet.Append(new OpenXml.CellFormats { Count = 0 });

            //add default styles
            AddFont(stylesheet, "Calibri", 11, false, null);
            AddFill(stylesheet, OpenXml.PatternValues.None, null, null);
            AddBorder(stylesheet, null, null, null, null, null);

            return stylesheet;
        }

        private static void AddFont(OpenXml.Stylesheet stylesheet, string name, double? size, bool isBold, OpenXml.Color color)
        {
            var id = stylesheet.Fonts.Count;
            var font = new OpenXml.Font();

            if (size.HasValue)
                font.Append(new OpenXml.FontSize { Val = size.Value });

            if (!string.IsNullOrEmpty(name))
                font.Append(new OpenXml.FontName { Val = name });

            if (isBold)
                font.Append(new OpenXml.Bold());

            if (color != null)
                font.Append(color);

            stylesheet.Fonts.Append(font);
            stylesheet.Fonts.Count++;
        }

        private static void AddFill(OpenXml.Stylesheet stylesheet, OpenXml.PatternValues type, OpenXml.Color fgColor, OpenXml.Color bgColor)
        {
            var id = stylesheet.Fills.Count;
            var fill = new OpenXml.Fill
            {
                PatternFill = new OpenXml.PatternFill { PatternType = type }
            };

            if (fgColor != null)
                fill.PatternFill.Append(new OpenXml.ForegroundColor { Rgb = fgColor.Rgb });

            if (bgColor != null)
                fill.PatternFill.Append(new OpenXml.BackgroundColor { Rgb = bgColor.Rgb });

            stylesheet.Fills.Append(fill);
            stylesheet.Fills.Count++;
        }

        private static void AddBorder(
            OpenXml.Stylesheet stylesheet,
            OpenXml.LeftBorder left,
            OpenXml.RightBorder right,
            OpenXml.TopBorder top,
            OpenXml.BottomBorder bottom,
            OpenXml.DiagonalBorder diagonal)
        {
            var id = stylesheet.Borders.Count;
            var border = new OpenXml.Border();

            border.LeftBorder = left ?? new OpenXml.LeftBorder();
            border.RightBorder = right ?? new OpenXml.RightBorder();
            border.TopBorder = top ?? new OpenXml.TopBorder();
            border.BottomBorder = bottom ?? new OpenXml.BottomBorder();
            border.DiagonalBorder = diagonal ?? new OpenXml.DiagonalBorder();

            stylesheet.Borders.Append(border);
            stylesheet.Borders.Count++;
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
