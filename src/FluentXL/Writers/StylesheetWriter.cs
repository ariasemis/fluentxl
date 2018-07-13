using DocumentFormat.OpenXml.Packaging;
using FluentXL.Models;
using FluentXL.Utils;
using System;
using System.Collections.Generic;
using System.Text;
using OpenXml = DocumentFormat.OpenXml.Spreadsheet;

namespace FluentXL.Writers
{
    internal class StylesheetWriter : IWriter<Stylesheet>
    {
        //TODO: current implementation uses the DOM API from OpenXML
        // change to use SAX for better performance

        private readonly WorkbookStylesPart workbookStylesPart;

        public StylesheetWriter(WorkbookStylesPart workbookStylesPart)
        {
            this.workbookStylesPart = workbookStylesPart;
        }

        public void Write(Stylesheet element)
        {
            // create stylesheet
            var stylesheet = new OpenXml.Stylesheet();

            Write(stylesheet, element.Fonts);
            Write(stylesheet, element.Fills);
            Write(stylesheet, element.Borders);
            Write(stylesheet, element.NumberFormats);
            Write(stylesheet, element.CellFormats);

            // write to part
            workbookStylesPart.Stylesheet = stylesheet;
            stylesheet.Save();
        }

        void Write(OpenXml.Stylesheet stylesheet, CountedCollection<Font> fonts)
        {
            if (fonts.Count > 0)
            {
                stylesheet.Append(new OpenXml.Fonts { Count = (uint)fonts.Count });

                foreach (var font in fonts)
                {
                    var f = new OpenXml.Font();

                    if (!string.IsNullOrEmpty(font.Name))
                        f.Append(new OpenXml.FontName { Val = font.Name });

                    if (font.Size.HasValue)
                        f.Append(new OpenXml.FontSize { Val = font.Size.Value });

                    if (font.Bold ?? false)
                        f.Append(new OpenXml.Bold());

                    stylesheet.Fonts.Append(f);
                }
            }
        }

        void Write(OpenXml.Stylesheet stylesheet, CountedCollection<Fill> fills)
        {
            if (fills.Count > 0)
            {
                stylesheet.Append(new OpenXml.Fills { Count = (uint)fills.Count });

                foreach (var fill in fills)
                {
                    var f = new OpenXml.Fill();

                    if (fill.PatternFill != null)
                    {
                        var rawType = (int)fill.PatternFill.PatternType;
                        var type = (OpenXml.PatternValues)rawType;

                        f.PatternFill = new OpenXml.PatternFill { PatternType = type };

                        if (fill.PatternFill.ForegroundColor != null)
                            f.PatternFill.Append(new OpenXml.ForegroundColor { Rgb = fill.PatternFill.ForegroundColor.Rgb });

                        if (fill.PatternFill.BackgroundColor != null)
                            f.PatternFill.Append(new OpenXml.BackgroundColor { Rgb = fill.PatternFill.BackgroundColor.Rgb });
                    }

                    stylesheet.Fills.Append(f);
                }
            }
        }

        void Write(OpenXml.Stylesheet stylesheet, CountedCollection<Border> borders)
        {
            if (borders.Count > 0)
            {
                stylesheet.Append(new OpenXml.Borders { Count = (uint)borders.Count });

                foreach (var border in borders)
                {
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

        void Write(OpenXml.Stylesheet stylesheet, CountedCollection<CellFormat> cellFormats)
        {
            if (cellFormats.Count > 0)
            {
                stylesheet.Append(new OpenXml.CellFormats { Count = (uint)cellFormats.Count });

                foreach (var cellFormat in cellFormats)
                {
                }
            }
        }

        void Write(OpenXml.Stylesheet stylesheet, CountedCollection<NumberFormat> numberFormats)
        {
            throw new NotImplementedException();
        }
    }
}
