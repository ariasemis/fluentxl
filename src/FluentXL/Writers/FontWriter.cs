using FluentXL.Models;
using FluentXL.Utils;
using System;
using OpenXml = DocumentFormat.OpenXml.Spreadsheet;

namespace FluentXL.Writers
{
    internal sealed class FontWriter :
        IWriter<Font>,
        IWriter<CountedCollection<Font>>
    {
        //TODO: current implementation uses the DOM API from OpenXML, change to use SAX for better performance

        private readonly OpenXml.Stylesheet stylesheet;

        public FontWriter(OpenXml.Stylesheet stylesheet)
        {
            this.stylesheet = stylesheet ?? throw new ArgumentNullException(nameof(stylesheet));
        }

        public void Write(CountedCollection<Font> fonts)
        {
            if (fonts == null)
                throw new ArgumentNullException(nameof(fonts));

            if (fonts.Count <= 0)
                return;

            stylesheet.Append(new OpenXml.Fonts { Count = (uint)fonts.Count });

            foreach (var font in fonts)
            {
                Write(font);
            }
        }

        public void Write(Font font)
        {
            if (font == null)
                throw new ArgumentNullException(nameof(font));

            var f = new OpenXml.Font();

            if (font.Size.HasValue)
                f.Append(new OpenXml.FontSize { Val = font.Size.Value });

            if (!string.IsNullOrEmpty(font.Name))
                f.Append(new OpenXml.FontName { Val = font.Name });

            if (font.Bold ?? false)
                f.Append(new OpenXml.Bold());

            stylesheet.Fonts.Append(f);
        }
    }
}
