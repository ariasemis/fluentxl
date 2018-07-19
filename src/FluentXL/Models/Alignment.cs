namespace FluentXL.Models
{
    public class Alignment
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
    }
}
