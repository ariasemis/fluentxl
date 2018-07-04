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
        IExpectBorderStylesheetSpecification WithOutline(BorderStyle style, IBuilderSpecification<Color> colorSpecification);
    }

    public interface IExpectBorderStylesheetSpecification
    {
        IBuilderSpecification<Border> OnStylesheet(IStylesheetSpecification stylesheet);
    }

    public interface IBorderSpecification : IExpectBorderSpecification, IExpectBorderStylesheetSpecification
    {
    }
}
