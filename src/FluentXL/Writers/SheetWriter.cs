using DocumentFormat.OpenXml;
using DocumentFormat.OpenXml.Packaging;
using FluentXL.Models;
using System;
using System.Collections.Generic;
using OpenXml = DocumentFormat.OpenXml.Spreadsheet;

namespace FluentXL.Writers
{
    internal class SheetWriter : IWriter<Sheet>
    {
        static readonly OpenXml.Worksheet sheetTemplate = new OpenXml.Worksheet();
        static readonly OpenXml.Columns columnsTemplate = new OpenXml.Columns();
        static readonly OpenXml.SheetData dataTemplate = new OpenXml.SheetData();
        static readonly OpenXml.MergeCells mergeCellsTemplate = new OpenXml.MergeCells();

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

            if (sheet.Columns != null) Write(sheet.Columns);
            if (sheet.Rows != null) Write(sheet.Rows);
            if (sheet.MergeCells != null) Write(sheet.MergeCells);

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

        void Write(IEnumerable<MergeCell> mergeCells)
        {
            //TODO
            //throw new NotImplementedException();
        }
    }
}
