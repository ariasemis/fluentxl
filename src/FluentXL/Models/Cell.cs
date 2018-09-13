using System;

namespace FluentXL.Models
{
    public class Cell
    {
        public Cell(
            uint row,
            uint column,
            CellType type,
            string value,
            uint? style = null)
        {
            if (row < Models.Row.MIN_INDEX)
                throw new ArgumentException($"The row value must be greater or equal to {Models.Row.MIN_INDEX}", nameof(row));
            if (row > Models.Row.MAX_INDEX)
                throw new ArgumentException($"The row value must be less or equal to {Models.Row.MAX_INDEX}", nameof(row));
            if (column < Models.Column.MIN_INDEX)
                throw new ArgumentException($"The column value must be greater or equal to {Models.Column.MIN_INDEX}", nameof(column));
            if (column > Models.Column.MAX_INDEX)
                throw new ArgumentException($"The column value must be less or equal to {Models.Column.MAX_INDEX}", nameof(column));

            Row = row;
            Column = column;
            Type = type;
            Value = value;
            Style = style;
        }

        public uint Row { get; }
        public uint Column { get; }
        public CellType Type { get; }
        public string Value { get; }
        public uint? Style { get; }
    }
}
