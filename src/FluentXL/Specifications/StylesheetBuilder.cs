using FluentXL.Models;
using FluentXL.Utils;
using System;
using System.Collections.Generic;

namespace FluentXL.Specifications
{
    internal class StylesheetBuilder : IStylesheetBuilder
    {
        //TODO: redesign to reuse styles when they have the same values

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

        #region Refactor

        public uint AddFont(Font font)
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

        public uint AddFill(Fill fill)
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

        public uint AddBorder(Border border)
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

        public uint AddCellFormat(CellFormat cellFormat)
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

        public uint AddNumberFormat(NumberFormat numberFormat)
        {
            if (numberFormat == null)
                throw new ArgumentNullException(nameof(numberFormat));

            var i = NumberFormats.IndexOf(numberFormat);

            if (i < 0)
            {
                i = NumberFormats.Count;
                NumberFormats.Add(numberFormat);
            }

            return (uint)i;
        }

        #endregion

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
            => NumberFormat.INITIAL_ID + (uint)NumberFormats.Count;

        private void AddDefaultStyles()
        {
            Fonts.Add(new Font(0, "Calibri", size: 11));
            Fills.Add(new Fill(0, new PatternFill(FillPattern.None, null, null), null));
            Fills.Add(new Fill(0, new PatternFill(FillPattern.Gray125, null, null), null));
            Borders.Add(new Border(0));
            CellFormats.Add(new CellFormat(0, null, 0, 0, 0, 0));
        }
    }
}
