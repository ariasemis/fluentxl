using FluentXL.Utils;
using System;
using System.Collections.Generic;

namespace FluentXL.Elements
{
    public class Sheet
    {
        public Sheet(
            string name,
            IEnumerable<Column> columns,
            IEnumerable<Row> rows,
            CountedCollection<MergeCell> mergeCells)
        {
            Name = name ?? throw new ArgumentNullException(nameof(name));
            Columns = columns ?? throw new ArgumentNullException(nameof(columns));
            Rows = rows ?? throw new ArgumentNullException(nameof(rows));
            MergeCells = mergeCells ?? throw new ArgumentNullException(nameof(mergeCells));
        }

        public string Name { get; }

        public IEnumerable<Column> Columns { get; }

        public IEnumerable<Row> Rows { get; }

        public CountedCollection<MergeCell> MergeCells { get; }
    }
}
