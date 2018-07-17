namespace FluentXL.Models
{
    public sealed class NumberFormat
    {
        public const uint INITIAL_ID = 164;

        public NumberFormat(
            uint id,
            string formatCode)
        {
            Id = id;
            FormatCode = formatCode;
        }

        public uint Id { get; }

        public string FormatCode { get; }
    }
}
