using FluentXL.Models;

namespace FluentXL.Specifications.Styles
{
    public interface IColorSpecification
    {
        IBuilderSpecification<Color> FromRgb(string value);
    }
}
