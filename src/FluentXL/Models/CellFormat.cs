namespace FluentXL.Models
{
    public class CellFormat
    {
        public CellFormat(
            uint id,
            uint? formatId = null,
            uint? fontId = null,
            uint? fillId = null,
            uint? borderId = null,
            uint? numberFormatId = null,
            bool? hasPivotButton = null,
            bool? hasQuotePrefix = null,
            Alignment alignment = null,
            Protection protection = null)
        {
            Id = id;
            FormatId = formatId;
            FontId = fontId;
            FillId = fillId;
            BorderId = borderId;
            NumberFormatId = numberFormatId;
            HasPivotButton = hasPivotButton;
            HasQuotePrefix = hasQuotePrefix;
            Alignment = alignment;
            Protection = protection;
        }

        public uint Id { get; }

        public uint? FormatId { get; }

        public uint? FontId { get; }

        public uint? FillId { get; }

        public uint? BorderId { get; }

        public uint? NumberFormatId { get; }

        public bool? HasPivotButton { get; }

        public bool? HasQuotePrefix { get; }

        public Alignment Alignment { get; }

        public Protection Protection { get; }
    }
}
