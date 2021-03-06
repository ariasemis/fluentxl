﻿using FluentXL.Elements;
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

            var cf = new OpenXml.CellFormat();

            if (cellFormat.FontId.HasValue)
            {
                cf.FontId = cellFormat.FontId.Value;
                cf.ApplyFont = cellFormat.FontId != 0;
            }

            if (cellFormat.FillId.HasValue)
            {
                cf.FillId = cellFormat.FillId.Value;
                cf.ApplyFill = cellFormat.FillId != 0;
            }

            if (cellFormat.BorderId.HasValue)
            {
                cf.BorderId = cellFormat.BorderId.Value;
                cf.ApplyBorder = cellFormat.BorderId != 0;
            }

            if (cellFormat.NumberFormatId.HasValue)
            {
                cf.NumberFormatId = cellFormat.NumberFormatId.Value;
                cf.ApplyNumberFormat = cellFormat.NumberFormatId != 0;
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
