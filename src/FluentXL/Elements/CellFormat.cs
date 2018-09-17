using System;

namespace FluentXL.Elements
{
    public class CellFormat : IEquatable<CellFormat>
    {
        public CellFormat(
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

        public uint? FormatId { get; }

        public uint? FontId { get; }

        public uint? FillId { get; }

        public uint? BorderId { get; }

        public uint? NumberFormatId { get; }

        public bool? HasPivotButton { get; }

        public bool? HasQuotePrefix { get; }

        public Alignment Alignment { get; }

        public Protection Protection { get; }

        #region Equality

        public static bool operator ==(CellFormat a, CellFormat b)
        {
            if (ReferenceEquals(a, b))
                return true;

            if (a is null || b is null)
                return false;

            return Equals(a, b);
        }

        public static bool operator !=(CellFormat a, CellFormat b)
            => !(a == b);

        public bool Equals(CellFormat other)
        {
            if (ReferenceEquals(this, other))
                return true;
            if (other is null)
                return false;

            return FormatId == other.FormatId
               && FontId == other.FontId
               && FillId == other.FillId
               && BorderId == other.BorderId
               && NumberFormatId == other.NumberFormatId
               && HasPivotButton == other.HasPivotButton
               && HasQuotePrefix == other.HasQuotePrefix
               && Alignment == other.Alignment
               && Protection == other.Protection;
        }

        public override bool Equals(object obj)
            => obj is CellFormat && Equals(obj as CellFormat);

        public override int GetHashCode()
        {
            unchecked
            {
                int hash = 11;

                hash = hash * 17 + FormatId?.GetHashCode() ?? 0;
                hash = hash * 17 + FontId?.GetHashCode() ?? 0;
                hash = hash * 17 + FillId?.GetHashCode() ?? 0;
                hash = hash * 17 + BorderId?.GetHashCode() ?? 0;
                hash = hash * 17 + NumberFormatId?.GetHashCode() ?? 0;
                hash = hash * 17 + HasPivotButton?.GetHashCode() ?? 0;
                hash = hash * 17 + HasQuotePrefix?.GetHashCode() ?? 0;
                hash = hash * 17 + Alignment?.GetHashCode() ?? 0;
                hash = hash * 17 + Protection?.GetHashCode() ?? 0;

                return hash;
            }
        }

        #endregion
    }
}
