using FluentXL.Models;

namespace FluentXL.Specifications.Styles
{
    public interface IColorSpecification
    {
        /// <summary>
        /// Specifies the color value.
        /// </summary>
        /// <param name="value">The RGB value</param>
        /// <returns>An instance of IBuilderSpecification<Color></returns>
        IBuilderSpecification<Color> FromRgb(string value);

        /// <summary>
        /// Specifies the color value.
        /// </summary>
        /// <param name="red">The red value</param>
        /// <param name="green">The green value</param>
        /// <param name="blue">The blue value</param>
        /// <returns>An instance of IBuilderSpecification<Color></returns>
        IBuilderSpecification<Color> FromRgb(byte red, byte green, byte blue);
    }
}
