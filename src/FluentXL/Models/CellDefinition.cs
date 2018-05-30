namespace FluentXL.Models
{
    public class CellDefinition
    {
        public CellDefinition(
            uint column,
            CellType type,
            string value,
            uint? style = null)
        {
            Column = column;
            Type = type;
            Value = value;
            Style = style;
        }

        public CellDefinition(
            CellDefinition definition)
        {
            Column = definition.Column;
            Type = definition.Type;
            Value = definition.Value;
            Style = definition.Style;
        }

        public uint Column { get; }
        public CellType Type { get; }
        public string Value { get; }
        public uint? Style { get; }
    }
}
