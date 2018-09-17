using FluentXL.Utils;
using System;

namespace FluentXL.Elements
{
    public sealed class Stylesheet
    {
        public Stylesheet(
            CountedCollection<Font> fonts,
            CountedCollection<Fill> fills,
            CountedCollection<Border> borders,
            CountedCollection<CellFormat> cellFormats,
            CountedCollection<NumberFormat> numberFormats)
        {
            Fonts = fonts ?? throw new ArgumentNullException(nameof(fonts));
            Fills = fills ?? throw new ArgumentNullException(nameof(fills));
            Borders = borders ?? throw new ArgumentNullException(nameof(borders));
            CellFormats = cellFormats ?? throw new ArgumentNullException(nameof(cellFormats));
            NumberFormats = numberFormats ?? throw new ArgumentNullException(nameof(numberFormats));
        }

        public CountedCollection<Font> Fonts { get; }

        public CountedCollection<Fill> Fills { get; }

        public CountedCollection<Border> Borders { get; }

        public CountedCollection<CellFormat> CellFormats { get; }

        public CountedCollection<NumberFormat> NumberFormats { get; }
    }
}
