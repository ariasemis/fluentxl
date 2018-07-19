using FluentXL.Models;
using FluentXL.Utils;
using System;
using OpenXml = DocumentFormat.OpenXml.Spreadsheet;

namespace FluentXL.Writers
{
    internal sealed class CellFormatWriter :
        IWriter<CellFormat>,
        IWriter<CountedCollection<CellFormat>>
    {
        //TODO: current implementation uses the DOM API from OpenXML, change to use SAX for better performance

        private readonly OpenXml.Stylesheet stylesheet;

        public CellFormatWriter(OpenXml.Stylesheet stylesheet)
        {
            this.stylesheet = stylesheet ?? throw new ArgumentNullException(nameof(stylesheet));
        }

        public void Write(CountedCollection<CellFormat> cellFormats)
        {
            if (cellFormats == null)
                throw new ArgumentNullException(nameof(cellFormats));

            if (cellFormats.Count <= 0)
                return;

            stylesheet.Append(new OpenXml.CellFormats { Count = (uint)cellFormats.Count });

            foreach (var cellFormat in cellFormats)
            {
                Write(cellFormat);
            }
        }

        public void Write(CellFormat cellFormat)
        {
            if (cellFormat == null)
                throw new ArgumentNullException(nameof(cellFormat));

            var cf = new OpenXml.CellFormat
            {
                NumberFormatId = 0,
                FontId = 0,
                FillId = 0,
                BorderId = 0,
                //FormatId = 0
            };

            if (cellFormat.FontId.HasValue)
            {
                cf.FontId = cellFormat.FontId.Value;
                cf.ApplyFont = true;
            }

            if (cellFormat.FillId.HasValue)
            {
                cf.FillId = cellFormat.FillId.Value;
                cf.ApplyFill = true;
            }

            if (cellFormat.BorderId.HasValue)
            {
                cf.BorderId = cellFormat.BorderId.Value;
                cf.ApplyBorder = true;
            }

            if (cellFormat.NumberFormatId.HasValue)
            {
                cf.NumberFormatId = cellFormat.NumberFormatId.Value;
                cf.ApplyNumberFormat = true;
            }

            if (cellFormat.Alignment != null)
            {
                IWriter<Alignment> alignmentWriter = new AlignmentWriter(cf);
                alignmentWriter.Write(cellFormat.Alignment);
            }

            stylesheet.CellFormats.Append(cf);
        }
    }
}
