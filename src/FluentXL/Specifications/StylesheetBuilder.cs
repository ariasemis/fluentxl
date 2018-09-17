using FluentXL.Elements;
using FluentXL.Utils;
using System;
using System.Collections.Generic;
using System.Linq;

namespace FluentXL.Specifications
{
    internal class StylesheetBuilder : IStylesheetBuilder
    {
        private IList<Font> Fonts { get; set; } = new List<Font>();
        private IList<Fill> Fills { get; set; } = new List<Fill>();
        private IList<Border> Borders { get; set; } = new List<Border>();
        private IList<CellFormat> CellFormats { get; set; } = new List<CellFormat>();
        private IList<NumberFormat> NumberFormats { get; set; } = new List<NumberFormat>();

        public StylesheetBuilder()
        {
            AddDefaultStyles();
        }

        public Stylesheet Build()
        {
            return new Stylesheet(
                new CountedCollection<Font>(Fonts.Count, Fonts),
                new CountedCollection<Fill>(Fills.Count, Fills),
                new CountedCollection<Border>(Borders.Count, Borders),
                new CountedCollection<CellFormat>(CellFormats.Count, CellFormats),
                new CountedCollection<NumberFormat>(NumberFormats.Count, NumberFormats));
        }

        public uint Add(Font font)
        {
            if (font == null)
                throw new ArgumentNullException(nameof(font));

            var i = Fonts.IndexOf(font);

            if (i < 0)
            {
                i = Fonts.Count;
                Fonts.Add(font);
            }

            return (uint)i;
        }

        public uint Add(Fill fill)
        {
            if (fill == null)
                throw new ArgumentNullException(nameof(fill));

            var i = Fills.IndexOf(fill);

            if (i < 0)
            {
                i = Fills.Count;
                Fills.Add(fill);
            }

            return (uint)i;
        }

        public uint Add(Border border)
        {
            if (border == null)
                throw new ArgumentNullException(nameof(border));

            var i = Borders.IndexOf(border);

            if (i < 0)
            {
                i = Borders.Count;
                Borders.Add(border);
            }

            return (uint)i;
        }

        public uint Add(CellFormat cellFormat)
        {
            if (cellFormat == null)
                throw new ArgumentNullException(nameof(cellFormat));

            var i = CellFormats.IndexOf(cellFormat);

            if (i < 0)
            {
                i = CellFormats.Count;
                CellFormats.Add(cellFormat);
            }

            return (uint)i;
        }

        public void Add(NumberFormat numberFormat)
        {
            if (numberFormat == null)
                throw new ArgumentNullException(nameof(numberFormat));
            if (numberFormat.Id != GetIndexForNumberFormat(numberFormat.FormatCode))
                throw new ArgumentException("The number format id does not match its current position in the stylesheet");

            if (!NumberFormats.Contains(numberFormat))
                NumberFormats.Add(numberFormat);
        }

        public uint GetIndexForNumberFormat(string formatCode)
        {
            if (formatCode == null)
                throw new ArgumentNullException(nameof(formatCode));

            var existing = NumberFormats.FirstOrDefault(x => x.FormatCode == formatCode);

            return (existing == null)
                ? NumberFormat.INITIAL_ID + (uint)NumberFormats.Count
                : existing.Id;
        }

        private void AddDefaultStyles()
        {
            Fonts.Add(new Font("Calibri", size: 11));
            Fills.Add(new Fill(new PatternFill(FillPattern.None, null, null), null));
            Fills.Add(new Fill(new PatternFill(FillPattern.Gray125, null, null), null));
            Borders.Add(new Border());
            CellFormats.Add(new CellFormat(null, 0, 0, 0, 0));
        }
    }
}
