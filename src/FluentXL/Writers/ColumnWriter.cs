using DocumentFormat.OpenXml;
using FluentXL.Elements;
using System;
using System.Collections.Generic;
using OpenXml = DocumentFormat.OpenXml.Spreadsheet;

namespace FluentXL.Writers
{
    internal class ColumnWriter : IWriter<Column>
    {
        static readonly OpenXml.Column template = new OpenXml.Column();
        readonly OpenXmlWriter writer;

        public ColumnWriter(OpenXmlWriter writer)
        {
            this.writer = writer;
        }

        public void Write(Column column)
        {
            if (column == null) throw new ArgumentNullException(nameof(column));

            writer.WriteStartElement(template, GetAttributes(column));
            writer.WriteEndElement();
        }

        static IEnumerable<OpenXmlAttribute> GetAttributes(Column column)
        {
            yield return new OpenXmlAttribute(Attributes.MIN, null, column.Index.ToString());
            yield return new OpenXmlAttribute(Attributes.MAX, null, column.Index.ToString());
            yield return new OpenXmlAttribute(Attributes.WIDTH, null, column.Width.ToString());
            yield return new OpenXmlAttribute(Attributes.CUSTOM_WIDTH, null, "1");

            yield break;
        }

        static class Attributes
        {
            public const string MIN = "min";
            public const string MAX = "max";
            public const string WIDTH = "width";
            public const string CUSTOM_WIDTH = "customWidth";
        }
    }
}
