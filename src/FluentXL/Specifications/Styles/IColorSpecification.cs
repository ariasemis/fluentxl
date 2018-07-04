using FluentXL.Models;

namespace FluentXL.Specifications.Styles
{
    public interface IColorSpecification
    {
        IBuilderSpecification<Color> FromRgb(string value);

        IBuilderSpecification<Color> FromRgb(byte red, byte green, byte blue);
    }
}
