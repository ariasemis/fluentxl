using FluentXL.Elements;
using FluentXL.Utils;
using System;
using OpenXml = DocumentFormat.OpenXml.Spreadsheet;

namespace FluentXL.Writers
{
    internal sealed class NumberFormatWriter :
        IWriter<NumberFormat>,
        IWriter<CountedCollection<NumberFormat>>
    {
        //TODO: current implementation uses the DOM API from OpenXML, change to use SAX for better performance

        private readonly OpenXml.Stylesheet stylesheet;

        public NumberFormatWriter(OpenXml.Stylesheet stylesheet)
        {
            this.stylesheet = stylesheet ?? throw new ArgumentNullException(nameof(stylesheet));
        }

        public void Write(CountedCollection<NumberFormat> numberFormats)
        {
            if (numberFormats == null)
                throw new ArgumentNullException(nameof(numberFormats));

            if (numberFormats.Count <= 0)
                return;

            //TODO: check this works properly
            stylesheet.NumberingFormats.Count = (uint)numberFormats.Count + NumberFormat.INITIAL_ID;

            foreach (var numberFormat in numberFormats)
            {
                Write(numberFormat);
            }
        }

        public void Write(NumberFormat numberFormat)
        {
            if (numberFormat == null)
                throw new ArgumentNullException(nameof(numberFormat));

            var nf = new OpenXml.NumberingFormat
            {
                NumberFormatId = numberFormat.Id,
                FormatCode = numberFormat.FormatCode
            };

            stylesheet.NumberingFormats.Append(nf);
        }
    }
}
