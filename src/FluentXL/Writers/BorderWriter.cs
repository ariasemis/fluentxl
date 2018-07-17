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

            var b = new OpenXml.Border
            {
                LeftBorder = new OpenXml.LeftBorder(),
                RightBorder = new OpenXml.RightBorder(),
                TopBorder = new OpenXml.TopBorder(),
                BottomBorder = new OpenXml.BottomBorder(),
                DiagonalBorder = new OpenXml.DiagonalBorder()
            };

            if (border.Left != null)
                SetBorderProperties(b.LeftBorder, border.Left);

            if (border.Right != null)
                SetBorderProperties(b.RightBorder, border.Right);

            if (border.Top != null)
                SetBorderProperties(b.TopBorder, border.Top);

            if (border.Bottom != null)
                SetBorderProperties(b.BottomBorder, border.Bottom);

            if (border.Diagonal != null)
            {
                SetBorderProperties(b.DiagonalBorder, border.Diagonal);

                b.DiagonalUp = border.Diagonal.Diagonal == BorderDiagonal.Up || border.Diagonal.Diagonal == BorderDiagonal.Both;
                b.DiagonalDown = border.Diagonal.Diagonal == BorderDiagonal.Down || border.Diagonal.Diagonal == BorderDiagonal.Both;
            }

            stylesheet.Borders.Append(b);
        }

        private void SetBorderProperties(OpenXml.BorderPropertiesType borderProperties, BorderSide borderSide)
        {
            borderProperties.Style = MapBorderStyle(borderSide.Style);

            if (borderSide.Color != null)
                borderProperties.Color = MapColor(borderSide.Color);
        }

        private static OpenXml.BorderStyleValues MapBorderStyle(BorderStyle style)
        {
            var raw = (int)style;
            return (OpenXml.BorderStyleValues)raw;
        }

        private static OpenXml.Color MapColor(Color color)
        {
            var c = new OpenXml.Color
            {
                Rgb = color.Rgb
            };

            if (color.Auto.HasValue)
                c.Auto = color.Auto.Value;
            if (color.Indexed.HasValue)
                c.Indexed = color.Indexed.Value;
            if (color.Theme.HasValue)
                c.Theme = color.Theme.Value;
            if (color.Tint.HasValue)
                c.Tint = color.Tint.Value;

            return c;
        }
    }
}
