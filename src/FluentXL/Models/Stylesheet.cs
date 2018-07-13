using System.Collections.Generic;

namespace FluentXL.Models
{
    public class Stylesheet
    {
        public IList<Font> Fonts { get; set; }

        public IList<Fill> Fills { get; set; }

        public IList<Border> Borders { get; set; }

        public IList<CellFormat> CellFormats { get; set; }

        public IList<NumberFormat> NumberFormats { get; set; }
    }
}
