using DocumentFormat.OpenXml;
using FluentXL.Elements;
using FluentXL.Utils;
using System;
using System.Collections.Generic;
using OpenXml = DocumentFormat.OpenXml.Spreadsheet;

namespace FluentXL.Writers
{
    internal class CellWriter : IWriter<Cell>
    {
        static readonly OpenXml.Cell templateCell = new OpenXml.Cell();
        static readonly OpenXml.CellValue templateCellValue = new OpenXml.CellValue();
        readonly OpenXmlWriter writer;

        public CellWriter(OpenXmlWriter writer)
        {
            this.writer = writer;
        }

        public void Write(Cell cell)
        {
            if (cell == null) throw new ArgumentNullException(nameof(cell));

            writer.WriteStartElement(templateCell, GetAttributes(cell));

            if (cell.Type == CellType.InlineString)
                WriteInlineString(cell.Value);
            else
                WriteCellValue(cell.Value);

            writer.WriteEndElement();
        }

        static IEnumerable<OpenXmlAttribute> GetAttributes(Cell cell)
        {
            yield return new OpenXmlAttribute(Attributes.REFERENCE, null, GetCellReference(cell));
            yield return new OpenXmlAttribute(Attributes.TYPE, null, GetCellType(cell.Type));
            if (cell.Style.HasValue) yield return new OpenXmlAttribute(Attributes.STYLE_INDEX, null, cell.Style.ToString());

            yield break;
        }

        static string GetCellReference(Cell cell)
            => ReferenceHelper.GetReference(cell.Row, cell.Column);

        static string GetCellType(CellType type)
        {
            switch (type)
            {
                case CellType.Boolean:
                    return Types.BOOLEAN;
                case CellType.Date:
                case CellType.Number:
                    return Types.NUMBER;
                case CellType.SharedString:
                    return Types.SHARED_STRING;
                case CellType.String:
                    return Types.STRING;
                case CellType.InlineString:
                    return Types.INLINE_STRING;
                default:
                    throw new ArgumentOutOfRangeException();
            }
        }

        void WriteCellValue(string value)
        {
            writer.WriteStartElement(templateCellValue);
            writer.WriteString(value);
            writer.WriteEndElement();
        }

        void WriteInlineString(string value)
        {
            writer.WriteElement(new OpenXml.InlineString(new OpenXml.Text(value)));
        }

        static class Attributes
        {
            public const string REFERENCE = "r";
            public const string STYLE_INDEX = "s";
            public const string TYPE = "t";
        }

        static class Types
        {
            public const string BOOLEAN = "b";
            public const string NUMBER = "n";
            public const string SHARED_STRING = "s";
            public const string STRING = "str";
            public const string INLINE_STRING = "inlineStr";
        }
    }
}
