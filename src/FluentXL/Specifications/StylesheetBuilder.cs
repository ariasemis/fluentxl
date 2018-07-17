using FluentXL.Models;
using FluentXL.Utils;
using System;
using System.Collections.Generic;

namespace FluentXL.Specifications
{
    internal class StylesheetBuilder : IStylesheetBuilder
    {
        private IList<Font> Fonts { get; set; } = new List<Font>();
        private IList<Fill> Fills { get; set; } = new List<Fill>();
        private IList<Border> Borders { get; set; } = new List<Border>();
        private IList<CellFormat> CellFormats { get; set; } = new List<CellFormat>();
        private IList<NumberFormat> NumberFormats { get; set; } = new List<NumberFormat>();

        public Stylesheet Build()
        {
            return new Stylesheet(
                new CountedCollection<Font>(Fonts.Count, Fonts),
                new CountedCollection<Fill>(Fills.Count, Fills),
                new CountedCollection<Border>(Borders.Count, Borders),
                new CountedCollection<CellFormat>(CellFormats.Count, CellFormats),
                new CountedCollection<NumberFormat>(NumberFormats.Count, NumberFormats));
        }

        public void Add(Font font)
        {
            if (font == null)
                throw new ArgumentNullException(nameof(font));
            if (font.Id != GenerateFontId())
                throw new ArgumentException("The font id does not match the current position in the stylesheet");

            Fonts.Add(font);
        }

        public void Add(Fill fill)
        {
            if (fill == null)
                throw new ArgumentNullException(nameof(fill));
            if (fill.Id != GenerateFillId())
                throw new ArgumentException("The fill id does not match the current position in the stylesheet");

            Fills.Add(fill);
        }

        public void Add(Border border)
        {
            if (border == null)
                throw new ArgumentNullException(nameof(border));
            if (border.Id != GenerateBorderId())
                throw new ArgumentException("The border id does not match the current position in the stylesheet");

            Borders.Add(border);
        }

        public void Add(CellFormat cellFormat)
        {
            if (cellFormat == null)
                throw new ArgumentNullException(nameof(cellFormat));
            if (cellFormat.Id != GenerateCellFormatId())
                throw new ArgumentException("The cell format id does not match the current position in the stylesheet");

            CellFormats.Add(cellFormat);
        }

        public void Add(NumberFormat numberFormat)
        {
            if (numberFormat == null)
                throw new ArgumentNullException(nameof(numberFormat));
            if (numberFormat.Id != GenerateNumberFormatId())
                throw new ArgumentException("The number format id does not match the current position in the stylesheet");

            NumberFormats.Add(numberFormat);
        }

        public uint GenerateBorderId()
            => (uint)Borders.Count;

        public uint GenerateCellFormatId()
            => (uint)CellFormats.Count;

        public uint GenerateFillId()
            => (uint)Fills.Count;

        public uint GenerateFontId()
            => (uint)Fonts.Count;

        public uint GenerateNumberFormatId()
            => NumberFormat.NUMBER_FORMAT_INITIAL_ID + (uint)NumberFormats.Count;
    }
}
