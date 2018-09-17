using FluentXL.Elements;

namespace FluentXL.Specifications.Styles
{
    public interface IExpectFillPattern
    {
        /// <summary>
        /// Specifies the fill pattern to be used.
        /// </summary>
        /// <param name="pattern">The fill pattern</param>
        /// <returns>An instance of IExpectFillColor</returns>
        IExpectFillColor WithPattern(FillPattern pattern);

        //WithGradient();
    }

    public interface IExpectFillColor
    {
        /// <summary>
        /// Specifies the background color.
        /// </summary>
        /// <param name="colorSpecification">A color specification</param>
        /// <returns>An instance of IFillSpecification</returns>
        IFillSpecification WithBackgroundColor(IBuilderSpecification<Color> colorSpecification);

        /// <summary>
        /// Specifies the foreground color.
        /// </summary>
        /// <param name="colorSpecification">A color specification</param>
        /// <returns>An instance of IFillSpecification</returns>
        IFillSpecification WithForegroundColor(IBuilderSpecification<Color> colorSpecification);
    }

    public interface IFillSpecification : IBuilderSpecification<Fill>, IExpectFillColor { }
}
