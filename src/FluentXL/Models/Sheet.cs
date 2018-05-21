using System.Collections.Generic;

namespace FluentXL.Models
{
    public class Sheet
    {
        public Sheet(string name, IEnumerable<Column> columns, IEnumerable<Row> rows, MergeCellCollection mergeCells)
        {
            Name = name;
            Columns = columns;
            Rows = rows;
            MergeCells = mergeCells;
        }

        public string Name { get; }
        public IEnumerable<Column> Columns { get; }
        public IEnumerable<Row> Rows { get; }
        public MergeCellCollection MergeCells { get; }
    }
}
