using FluentXL.Models;

namespace FluentXL.Specifications.CellFormats
{
    public interface IAlignmentSpecification
    {
        IAlignmentSpecification WithHorizontal(HorizontalAlignment horizontalAlignment);

        IAlignmentSpecification WithVertical(VerticalAlignment verticalAlignment);

        IAlignmentSpecification WithOrientation(AlignmentOrientation orientation);

        IAlignmentSpecification ShrinkToFit();

        IAlignmentSpecification WrapText();
    }
}
