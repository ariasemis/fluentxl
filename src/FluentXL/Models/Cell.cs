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
