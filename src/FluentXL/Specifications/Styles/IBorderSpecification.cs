using FluentXL.Elements;

namespace FluentXL.Specifications.Styles
{
    public interface IExpectBorderSpecification
    {
        /// <summary>
        /// Specifies the top of the border.
        /// </summary>
        /// <param name="style">A border style</param>
        /// <param name="colorSpecification">A color specification</param>
        /// <returns>An instance of IBorderSpecification</returns>
        IBorderSpecification WithTop(BorderStyle style, IBuilderSpecification<Color> colorSpecification);

        /// <summary>
        /// Specifies the bottom of the border.
        /// </summary>
        /// <param name="style">A border style</param>
        /// <param name="colorSpecification">A color specification</param>
        /// <returns>An instance of IBorderSpecification</returns>
        IBorderSpecification WithBottom(BorderStyle style, IBuilderSpecification<Color> colorSpecification);

        /// <summary>
        /// Specifies the left of the border.
        /// </summary>
        /// <param name="style">A border style</param>
        /// <param name="colorSpecification">A color specification</param>
        /// <returns>An instance of IBorderSpecification</returns>
        IBorderSpecification WithLeft(BorderStyle style, IBuilderSpecification<Color> colorSpecification);

        /// <summary>
        /// Specifies the right of the border.
        /// </summary>
        /// <param name="style">A border style</param>
        /// <param name="colorSpecification">A color specification</param>
        /// <returns>An instance of IBorderSpecification</returns>
        IBorderSpecification WithRight(BorderStyle style, IBuilderSpecification<Color> colorSpecification);

        /// <summary>
        /// Specifies the diagonal of the border.
        /// </summary>
        /// <param name="style">A border style</param>
        /// <param name="colorSpecification">A color specification</param>
        /// <param name="diagonal">The diagonals affected</param>
        /// <returns>An instance of IBorderSpecification</returns>
        IBorderSpecification WithDiagonal(BorderStyle style, IBuilderSpecification<Color> colorSpecification, BorderDiagonal diagonal);
    }

    public interface IExpectBorderOutlineSpecification : IExpectBorderSpecification
    {
        /// <summary>
        /// Specifies the outline of the border.
        /// </summary>
        /// <param name="style">A border style</param>
        /// <param name="colorSpecification">A color specification</param>
        /// <returns>An instance of IBuilderSpecification<Border></returns>
        IBuilderSpecification<Border> WithOutline(BorderStyle style, IBuilderSpecification<Color> colorSpecification);
    }

    public interface IBorderSpecification : IExpectBorderSpecification, IBuilderSpecification<Border>
    {
    }
}
