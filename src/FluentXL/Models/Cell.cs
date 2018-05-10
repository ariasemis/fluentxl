namespace FluentXL.Models
{
    public class Cell
    {
        public uint Row { get; }
        public uint Column { get; }
        public CellType DataType { get; }
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
