using FluentXL.Models;

namespace FluentXL.Specifications.Styles
{
    public class AlignmentSpecification : IAlignmentSpecification
    {
        private HorizontalAlignment Horizontal { get; set; } = HorizontalAlignment.Center;
        private VerticalAlignment Vertical { get; set; } = VerticalAlignment.Top;
        private bool Wrap { get; set; } = true;
        private ReadingOrder ReadingOrder { get; set; } = ReadingOrder.LeftToRight;
        private uint? TextRotation { get; set; } = null;
        private uint? Indent { get; set; } = null;
        private uint? RelativeIndent { get; set; } = null;
        private bool? JustifyLastLine { get; set; } = null;
        private bool? Shrink { get; set; } = null;
        private string MergeCell { get; set; } = null;

        private AlignmentSpecification() { }

        private AlignmentSpecification(AlignmentSpecification specification)
        {
            Horizontal = specification.Horizontal;
            Vertical = specification.Vertical;
            Wrap = specification.Wrap;
            ReadingOrder = specification.ReadingOrder;
            TextRotation = specification.TextRotation;
            Indent = specification.Indent;
            RelativeIndent = specification.RelativeIndent;
            JustifyLastLine = specification.JustifyLastLine;
            Shrink = specification.Shrink;
            MergeCell = specification.MergeCell;
        }

        public static IAlignmentSpecification New()
            => new AlignmentSpecification();

        public Alignment Build()
        {
            return new Alignment(
                Horizontal,
                Vertical,
                Wrap,
                ReadingOrder,
                TextRotation,
                Indent,
                RelativeIndent,
                JustifyLastLine,
                Shrink,
                MergeCell);
        }

        public IAlignmentSpecification ShrinkToFit()
        {
            return new AlignmentSpecification(this)
            {
                Shrink = true
            };
        }

        public IAlignmentSpecification WithHorizontal(HorizontalAlignment horizontalAlignment)
        {
            return new AlignmentSpecification(this)
            {
                Horizontal = horizontalAlignment
            };
        }

        public IAlignmentSpecification WithReadingOrder(ReadingOrder readingOrder)
        {
            return new AlignmentSpecification(this)
            {
                ReadingOrder = readingOrder
            };
        }

        public IAlignmentSpecification WithVertical(VerticalAlignment verticalAlignment)
        {
            return new AlignmentSpecification(this)
            {
                Vertical = verticalAlignment
            };
        }

        public IAlignmentSpecification WrapText()
        {
            return new AlignmentSpecification(this)
            {
                Wrap = true
            };
        }
    }
}
