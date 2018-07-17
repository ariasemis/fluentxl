using FluentXL.Models;
using FluentXL.Utils;
using System;
using OpenXml = DocumentFormat.OpenXml.Spreadsheet;

namespace FluentXL.Writers
{
    internal sealed class BorderWriter :
        IWriter<Border>,
        IWriter<CountedCollection<Border>>
    {
        //TODO: current implementation uses the DOM API from OpenXML, change to use SAX for better performance

        private readonly OpenXml.Stylesheet stylesheet;

        public BorderWriter(OpenXml.Stylesheet stylesheet)
        {
            this.stylesheet = stylesheet ?? throw new ArgumentNullException(nameof(stylesheet));
        }

        public void Write(CountedCollection<Border> borders)
        {
            if (borders == null)
                throw new ArgumentNullException(nameof(borders));

            if (borders.Count <= 0)
                return;

            stylesheet.Append(new OpenXml.Borders { Count = (uint)borders.Count });

            foreach (var border in borders)
            {
                Write(border);
            }
        }

        public void Write(Border border)
        {
            if (border == null)
                throw new ArgumentNullException(nameof(border));

            var b = new OpenXml.Border();

            b.LeftBorder = new OpenXml.LeftBorder();
            if (border.Left != null)
            {
                //TODO: set properties
            }

            b.RightBorder = new OpenXml.RightBorder();
            if (border.Right != null)
            {
                //TODO: set properties
            }

            b.TopBorder = new OpenXml.TopBorder();
            if (border.Top != null)
            {
                //TODO: set properties
            }

            b.BottomBorder = new OpenXml.BottomBorder();
            if (border.Bottom != null)
            {
                //TODO: set properties
            }

            b.DiagonalBorder = new OpenXml.DiagonalBorder();
            if (border.Diagonal != null)
            {
                //TODO: set properties
            }

            stylesheet.Borders.Append(b);
        }
    }
}
