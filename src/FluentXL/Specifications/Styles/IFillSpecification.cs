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
        IExpectFillColor WithBackgroundColor(IBuilderSpecification<Color> colorSpecification);

        IExpectFillColor WithForegroundColor(IBuilderSpecification<Color> colorSpecification);

    }
}
