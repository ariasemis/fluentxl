using FluentXL.Models;

namespace FluentXL.Specifications.CellFormats
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
