using FluentXL.Models;

namespace FluentXL.Specifications.CellFormats
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
