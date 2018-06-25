using FluentXL.Models;

namespace FluentXL.Specifications.CellFormats
{
    public interface IBorderSpecification
    {
        IBorderSpecification WithTop(BorderStyle style, IBuilderSpecification<Color> colorSpecification);

        IBorderSpecification WithBottom(BorderStyle style, IBuilderSpecification<Color> colorSpecification);

        IBorderSpecification WithLeft(BorderStyle style, IBuilderSpecification<Color> colorSpecification);

        IBorderSpecification WithRight(BorderStyle style, IBuilderSpecification<Color> colorSpecification);

        IBorderSpecification WithDiagonal(BorderStyle style, IBuilderSpecification<Color> colorSpecification, BorderDiagonal diagonal);
    }
}
