using DocumentFormat.OpenXml;
using DocumentFormat.OpenXml.Packaging;
using FluentXL.Models;
using System;
using System.Collections.Generic;
using System.Linq;
using OpenXml = DocumentFormat.OpenXml.Spreadsheet;

namespace FluentXL.Writers
{
    internal class SheetWriter : IWriter<Sheet>
    {
        static readonly OpenXml.Worksheet sheetTemplate = new OpenXml.Worksheet();
        static readonly OpenXml.Columns columnsTemplate = new OpenXml.Columns();
        static readonly OpenXml.SheetData dataTemplate = new OpenXml.SheetData();

        readonly OpenXmlWriter writer;

        public SheetWriter(WorksheetPart worksheetPart)
        {
            writer = OpenXmlWriter.Create(worksheetPart);
        }

        public void Write(Sheet sheet)
        {
            if (sheet == null)
                throw new ArgumentNullException(nameof(sheet));

            writer.WriteStartElement(sheetTemplate);

            if (sheet.Columns?.Any() ?? false) Write(sheet.Columns);
            if (sheet.Rows?.Any() ?? false) Write(sheet.Rows);
            if (sheet.MergeCells?.Any() ?? false) Write(sheet.MergeCells);

            writer.WriteEndElement();
            writer.Close();
        }

        void Write(IEnumerable<Column> columns)
        {
            writer.WriteStartElement(columnsTemplate);

            IWriter<Column> columnWriter = new ColumnWriter(writer);
            foreach (var column in columns)
            {
                columnWriter.Write(column);
            }

            writer.WriteEndElement();
        }

        void Write(IEnumerable<Row> rows)
        {
            writer.WriteStartElement(dataTemplate);

            IWriter<Row> rowWriter = new RowWriter(writer);
            foreach (var row in rows)
            {
                rowWriter.Write(row);
            }

            writer.WriteEndElement();
        }

        void Write(MergeCellCollection mergeCells)
        {
            IWriter<MergeCellCollection> mergeCellWriter = new MergeCellWriter(writer);
            mergeCellWriter.Write(mergeCells);
        }
    }
}
