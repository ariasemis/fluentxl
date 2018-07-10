using FluentXL.Models;

namespace FluentXL.Specifications.Styles
{
    public interface IExpectFillPattern
    {
        IExpectFillColor WithPattern(FillPattern pattern);

        //WithGradient();
    }

    public interface IExpectFillColor
    {
        IFillSpecification WithBackgroundColor(IBuilderSpecification<Color> colorSpecification);

        IFillSpecification WithForegroundColor(IBuilderSpecification<Color> colorSpecification);
    }

    public interface IFillSpecification : IBuilderSpecification<Fill>, IExpectFillColor { }
}
