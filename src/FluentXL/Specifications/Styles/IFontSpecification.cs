using FluentXL.Elements;

namespace FluentXL.Specifications.Styles
{
    public interface IExpectFontName
    {
        /// <summary>
        /// Specifies the font name to be used.
        /// </summary>
        /// <param name="name">The name of the font</param>
        /// <returns>An instance of IExpectFontOptions</returns>
        IExpectFontOptions WithFont(string name);
    }

    public interface IExpectFontOptions : IBuilderSpecification<Font>
    {
        /// <summary>
        /// Specifies the size of the font.
        /// </summary>
        /// <param name="size">The size of the font</param>
        /// <returns>An instance of IExpectFontOptions</returns>
        IExpectFontOptions WithSize(double size);

        /// <summary>
        /// Specifies the color of the font.
        /// </summary>
        /// <param name="colorSpecification">A color specification</param>
        /// <returns>An instance of IExpectFontOptions</returns>
        IExpectFontOptions WithColor(IBuilderSpecification<Color> colorSpecification);

        /// <summary>
        /// Indicates that the font must be bold.
        /// </summary>
        /// <returns>An instance of IExpectFontOptions</returns>
        IExpectFontOptions Bold();

        /// <summary>
        /// Indicates that the font must be italic.
        /// </summary>
        /// <returns>An instance of IExpectFontOptions</returns>
        IExpectFontOptions Italic();
    }
}
