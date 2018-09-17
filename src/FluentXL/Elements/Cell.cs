using System;

namespace FluentXL.Elements
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
            if (row < Elements.Row.MIN_INDEX)
                throw new ArgumentException($"The row value must be greater or equal to {Elements.Row.MIN_INDEX}", nameof(row));
            if (row > Elements.Row.MAX_INDEX)
                throw new ArgumentException($"The row value must be less or equal to {Elements.Row.MAX_INDEX}", nameof(row));
            if (column < Elements.Column.MIN_INDEX)
                throw new ArgumentException($"The column value must be greater or equal to {Elements.Column.MIN_INDEX}", nameof(column));
            if (column > Elements.Column.MAX_INDEX)
                throw new ArgumentException($"The column value must be less or equal to {Elements.Column.MAX_INDEX}", nameof(column));

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
