using System;
using System.Collections.Generic;
using System.Text;

namespace FluentXL.Models
{
    class Stylesheet
    {
        public IList<Font> Fonts { get; set; }
        public IList<Fill> Fills { get; set; }
        public IList<Border> Borders { get; set; }

    }
}
