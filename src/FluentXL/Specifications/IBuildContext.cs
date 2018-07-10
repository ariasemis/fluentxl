using FluentXL.Specifications.Styles;

namespace FluentXL.Specifications
{
    public interface IBuildContext
    {
        IStylesheetSpecification Stylesheet { get; }
    }
}
