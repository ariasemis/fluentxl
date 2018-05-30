namespace FluentXL.Models
{
    public class Cell : CellDefinition
    {
        public Cell(
            uint row,
            uint column,
            CellType type,
            string value,
            uint? style = null)
            : base(column, type, value, style)
        {
            Row = row;
        }

        public Cell(
            uint row,
            CellDefinition definition)
            : base(definition)
        {
            Row = row;
        }

        public uint Row { get; }
    }
}
