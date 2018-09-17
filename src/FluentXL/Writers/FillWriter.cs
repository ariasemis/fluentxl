using FluentXL.Elements;
using FluentXL.Utils;
using System;
using OpenXml = DocumentFormat.OpenXml.Spreadsheet;

namespace FluentXL.Writers
{
    internal sealed class FillWriter :
        IWriter<Fill>,
        IWriter<CountedCollection<Fill>>
    {
        //TODO: current implementation uses the DOM API from OpenXML, change to use SAX for better performance

        private readonly OpenXml.Stylesheet stylesheet;

        public FillWriter(OpenXml.Stylesheet stylesheet)
        {
            this.stylesheet = stylesheet ?? throw new ArgumentNullException(nameof(stylesheet));
        }

        public void Write(CountedCollection<Fill> fills)
        {
            if (fills == null)
                throw new ArgumentNullException(nameof(fills));

            if (fills.Count <= 0)
                return;

            stylesheet.Append(new OpenXml.Fills { Count = (uint)fills.Count });

            foreach (var fill in fills)
            {
                Write(fill);
            }
        }

        public void Write(Fill fill)
        {
            if (fill == null)
                throw new ArgumentNullException(nameof(fill));

            var f = new OpenXml.Fill();

            if (fill.PatternFill != null)
            {
                f.PatternFill = new OpenXml.PatternFill { PatternType = MapFillPattern(fill.PatternFill.PatternType) };

                if (fill.PatternFill.ForegroundColor != null)
                    f.PatternFill.Append(fill.PatternFill.ForegroundColor.MapToForegroundColor());

                if (fill.PatternFill.BackgroundColor != null)
                    f.PatternFill.Append(fill.PatternFill.BackgroundColor.MapToBackgroundColor());
            }

            stylesheet.Fills.Append(f);
        }

        private static OpenXml.PatternValues MapFillPattern(FillPattern fillPattern)
        {
            var rawType = (int)fillPattern;
            return (OpenXml.PatternValues)rawType;
        }
    }
}
