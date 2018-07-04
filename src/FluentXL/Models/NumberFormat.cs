namespace FluentXL.Models
{
    public sealed class NumberFormat
    {
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
