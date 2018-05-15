using DocumentFormat.OpenXml;
using FluentXL.Models;
using System;
using System.Collections.Generic;
using OpenXml = DocumentFormat.OpenXml.Spreadsheet;

namespace FluentXL.Writers
{
    internal class RowWriter : IWriter<Row>
    {
        static readonly OpenXml.Row template = new OpenXml.Row();
        readonly OpenXmlWriter writer;

        public RowWriter(OpenXmlWriter writer)
        {
            this.writer = writer;
        }

        public void Write(Row row)
        {
            if (row == null) throw new ArgumentNullException(nameof(row));

            writer.WriteStartElement(template, GetAttributes(row));

            if (row.Cells != null)
            {
                IWriter<Cell> cellWriter = new CellWriter(writer);
                foreach (var cell in row.Cells)
                {
                    cellWriter.Write(cell);
                }
            }

            writer.WriteEndElement();
        }

        static IEnumerable<OpenXmlAttribute> GetAttributes(Row row)
        {
            yield return new OpenXmlAttribute(Attributes.INDEX, null, row.Index.ToString());
            if (row.Height.HasValue)
            {
                yield return new OpenXmlAttribute(Attributes.HEIGHT, null, row.Height.ToString());
                yield return new OpenXmlAttribute(Attributes.CUSTOM_HEIGHT, null, "1");
            }
            if (row.Style.HasValue)
            {
                yield return new OpenXmlAttribute(Attributes.STYLE_INDEX, null, row.Style.Value.ToString());
                yield return new OpenXmlAttribute(Attributes.CUSTOM_FORMAT, null, "1");
            }
            if (row.OutlineLevel.HasValue) yield return new OpenXmlAttribute(Attributes.OUTLINE_LEVEL, null, row.OutlineLevel.ToString());
            if (row.Collapsed.HasValue) yield return new OpenXmlAttribute(Attributes.COLLAPSED, null, row.Collapsed.Value ? "1" : "0");
            if (row.Hidden.HasValue) yield return new OpenXmlAttribute(Attributes.HIDDEN, null, row.Hidden.Value ? "1" : "0");
            if (row.ThickBottom.HasValue) yield return new OpenXmlAttribute(Attributes.THICK_BOTTOM, null, row.ThickBottom.Value ? "1" : "0");
            if (row.ThickTop.HasValue) yield return new OpenXmlAttribute(Attributes.THICK_TOP, null, row.ThickTop.Value ? "1" : "0");

            yield break;
        }

        static class Attributes
        {
            public const string INDEX = "r";
            public const string COLLAPSED = "collapsed";
            public const string CUSTOM_FORMAT = "customFormat";
            public const string CUSTOM_HEIGHT = "customHeight";
            public const string HEIGHT = "ht";
            public const string HIDDEN = "hidden";
            public const string OUTLINE_LEVEL = "outlineLevel";
            public const string STYLE_INDEX = "s";
            public const string THICK_BOTTOM = "thickBot";
            public const string THICK_TOP = "thickTop";
        }
    }
}
