using FluentXL.Utils;

namespace FluentXL.Models
{
    public class Stylesheet
    {
        public Stylesheet(
            CountedCollection<Font> fonts,
            CountedCollection<Fill> fills,
            CountedCollection<Border> borders,
            CountedCollection<CellFormat> cellFormats,
            CountedCollection<NumberFormat> numberFormats)
        {
            Fonts = fonts;
            Fills = fills;
            Borders = borders;
            CellFormats = cellFormats;
            NumberFormats = numberFormats;
        }

        public CountedCollection<Font> Fonts { get; }

        public CountedCollection<Fill> Fills { get; }

        public CountedCollection<Border> Borders { get; }

        public CountedCollection<CellFormat> CellFormats { get; }

        public CountedCollection<NumberFormat> NumberFormats { get; }
    }
}
