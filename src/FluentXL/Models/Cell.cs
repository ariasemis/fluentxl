namespace FluentXL.Models
{
    public class Cell
    {
        public Cell(
            uint row,
            uint column,
            CellType cellType,
            string value,
            uint? style = null)
        {
            Row = row;
            Column = column;
            CellType = cellType;
            Value = value;
            Style = style;
        }

        public uint Row { get; }
        public uint Column { get; }
        public CellType CellType { get; }
        public string Value { get; }
        public uint? Style { get; }
    }

    public enum CellType
    {
        Boolean,
        Number,
        SharedString,
        String,
        InlineString,
        Date
    }
}
