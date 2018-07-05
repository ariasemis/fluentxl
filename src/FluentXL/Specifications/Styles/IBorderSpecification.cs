using FluentXL.Models;

namespace FluentXL.Specifications.Styles
{
    public interface IExpectBorderSpecification
    {
        IBorderSpecification WithTop(BorderStyle style, IBuilderSpecification<Color> colorSpecification);

        IBorderSpecification WithBottom(BorderStyle style, IBuilderSpecification<Color> colorSpecification);

        IBorderSpecification WithLeft(BorderStyle style, IBuilderSpecification<Color> colorSpecification);

        IBorderSpecification WithRight(BorderStyle style, IBuilderSpecification<Color> colorSpecification);

        IBorderSpecification WithDiagonal(BorderStyle style, IBuilderSpecification<Color> colorSpecification, BorderDiagonal diagonal);
    }

    public interface IExpectBorderOutlineSpecification : IExpectBorderSpecification
    {
        IBuilderSpecification<Border> WithOutline(BorderStyle style, IBuilderSpecification<Color> colorSpecification);
    }

    public interface IBorderSpecification : IExpectBorderSpecification, IBuilderSpecification<Border>
    {
    }
}
