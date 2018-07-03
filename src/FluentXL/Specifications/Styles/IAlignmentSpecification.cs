using FluentXL.Models;

namespace FluentXL.Specifications.Styles
{
    public interface IAlignmentSpecification : IBuilderSpecification<Alignment>
    {
        IAlignmentSpecification WithHorizontal(HorizontalAlignment horizontalAlignment);

        IAlignmentSpecification WithVertical(VerticalAlignment verticalAlignment);

        IAlignmentSpecification WithReadingOrder(ReadingOrder readingOrder);

        IAlignmentSpecification ShrinkToFit();

        IAlignmentSpecification WrapText();
    }
}
