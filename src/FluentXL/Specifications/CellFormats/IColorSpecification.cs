using FluentXL.Models;

namespace FluentXL.Specifications.CellFormats
{
    public interface IColorSpecification
    {
        IBuilderSpecification<Color> FromRgb(string value);
    }
}
