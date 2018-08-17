using System;

namespace FluentXL.Models
{
    public class Alignment : IEquatable<Alignment>
    {
        public Alignment(
            HorizontalAlignment horizontal,
            VerticalAlignment vertical,
            bool wrapText,
            ReadingOrder readingOrder,
            uint? textRotation = null,
            uint? indent = null,
            int? relativeIndent = null,
            bool? justifyLastLine = null,
            bool? shrinkToFit = null,
            string mergeCell = null)
        {
            Horizontal = horizontal;
            Vertical = vertical;
            Wrap = wrapText;
            ReadingOrder = readingOrder;
            TextRotation = textRotation;
            Indent = indent;
            RelativeIndent = relativeIndent;
            JustifyLastLine = justifyLastLine;
            Shrink = shrinkToFit;
            MergeCell = mergeCell;
        }

        public HorizontalAlignment Horizontal { get; }

        public VerticalAlignment Vertical { get; }

        public bool Wrap { get; }

        public ReadingOrder ReadingOrder { get; }

        public uint? TextRotation { get; }

        public uint? Indent { get; }

        public int? RelativeIndent { get; }

        public bool? JustifyLastLine { get; }

        public bool? Shrink { get; }

        public string MergeCell { get; }

        public static bool operator ==(Alignment a, Alignment b)
        {
            if (ReferenceEquals(a, b))
                return true;

            if (a is null || b is null)
                return false;

            return Equals(a, b);
        }

        public static bool operator !=(Alignment a, Alignment b)
            => !(a == b);

        public bool Equals(Alignment other)
        {
            if (ReferenceEquals(this, other))
                return true;
            if (other is null)
                return false;

            return Horizontal == other.Horizontal
                && Vertical == other.Vertical
                && Wrap == other.Wrap
                && ReadingOrder == other.ReadingOrder
                && TextRotation == other.TextRotation
                && Indent == other.Indent
                && RelativeIndent == other.RelativeIndent
                && JustifyLastLine == other.JustifyLastLine
                && Shrink == other.Shrink
                && MergeCell == other.MergeCell;
        }

        public override bool Equals(object obj)
            => obj is Alignment && Equals(obj as Alignment);

        public override int GetHashCode()
        {
            unchecked
            {
                int hash = 19;

                hash = hash * 29 + Horizontal.GetHashCode();
                hash = hash * 29 + Vertical.GetHashCode();
                hash = hash * 29 + Wrap.GetHashCode();
                hash = hash * 29 + ReadingOrder.GetHashCode();
                hash = hash * 29 + TextRotation?.GetHashCode() ?? 0;
                hash = hash * 29 + Indent?.GetHashCode() ?? 0;
                hash = hash * 29 + RelativeIndent?.GetHashCode() ?? 0;
                hash = hash * 29 + JustifyLastLine?.GetHashCode() ?? 0;
                hash = hash * 29 + Shrink?.GetHashCode() ?? 0;
                hash = hash * 29 + MergeCell?.GetHashCode() ?? 0;

                return hash;
            }
        }
    }
}
