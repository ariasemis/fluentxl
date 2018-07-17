﻿using FluentXL.Models;
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