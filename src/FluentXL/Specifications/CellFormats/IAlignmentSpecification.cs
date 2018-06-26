using FluentXL.Models;

namespace FluentXL.Specifications.CellFormats
{
    public interface IAlignmentSpecification
    {
        IBuildingAlignmentSpecification WithHorizontal(HorizontalAlignment horizontalAlignment);

        IBuildingAlignmentSpecification WithVertical(VerticalAlignment verticalAlignment);

        IBuildingAlignmentSpecification WithOrientation(AlignmentOrientation orientation);

        IBuildingAlignmentSpecification ShrinkToFit();

        IBuildingAlignmentSpecification WrapText();
    }

    public interface IBuildingAlignmentSpecification : IBuilderSpecification<Alignment>, IAlignmentSpecification
    {
    }
}
