using DocumentFormat.OpenXml.Packaging;
using FluentXL.Elements;
using FluentXL.Utils;
using System;
using OpenXml = DocumentFormat.OpenXml.Spreadsheet;

namespace FluentXL.Writers
{
    internal sealed class StylesheetWriter : IWriter<Stylesheet>
    {
        //TODO: current implementation uses the DOM API from OpenXML, change to use SAX for better performance

        private readonly WorkbookStylesPart workbookStylesPart;

        public StylesheetWriter(WorkbookStylesPart workbookStylesPart)
        {
            this.workbookStylesPart = workbookStylesPart ?? throw new ArgumentNullException(nameof(workbookStylesPart));
        }

        public void Write(Stylesheet element)
        {
            if (element == null)
                throw new ArgumentNullException(nameof(element));

            // create stylesheet
            var stylesheet = new OpenXml.Stylesheet();

            Write(stylesheet, element.Fonts);
            Write(stylesheet, element.Fills);
            Write(stylesheet, element.Borders);
            Write(stylesheet, element.NumberFormats);
            Write(stylesheet, element.CellFormats);

            // write to part
            workbookStylesPart.Stylesheet = stylesheet;
            stylesheet.Save();
        }

        void Write(OpenXml.Stylesheet stylesheet, CountedCollection<Font> fonts)
        {
            IWriter<CountedCollection<Font>> fontWriter = new FontWriter(stylesheet);
            fontWriter.Write(fonts);
        }

        void Write(OpenXml.Stylesheet stylesheet, CountedCollection<Fill> fills)
        {
            IWriter<CountedCollection<Fill>> fillWriter = new FillWriter(stylesheet);
            fillWriter.Write(fills);
        }

        void Write(OpenXml.Stylesheet stylesheet, CountedCollection<Border> borders)
        {
            IWriter<CountedCollection<Border>> borderWriter = new BorderWriter(stylesheet);
            borderWriter.Write(borders);
        }

        void Write(OpenXml.Stylesheet stylesheet, CountedCollection<CellFormat> cellFormats)
        {
            IWriter<CountedCollection<CellFormat>> cellFormatWriter = new CellFormatWriter(stylesheet);
            cellFormatWriter.Write(cellFormats);
        }

        void Write(OpenXml.Stylesheet stylesheet, CountedCollection<NumberFormat> numberFormats)
        {
            IWriter<CountedCollection<NumberFormat>> numberFormatWriter = new NumberFormatWriter(stylesheet);
            numberFormatWriter.Write(numberFormats);
        }
    }
}
