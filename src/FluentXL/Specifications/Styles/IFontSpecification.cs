using FluentXL.Models;

namespace FluentXL.Specifications.Styles
{
    public interface IExpectFontName
    {
        IExpectFontOptions WithFont(string name);
    }

    public interface IExpectFontOptions : IBuilderSpecification<Font>
    {
        IExpectFontOptions WithSize(double size);

        IExpectFontOptions WithColor(IBuilderSpecification<Color> colorSpecification);

        IExpectFontOptions Bold();

        IExpectFontOptions Italic();
    }
}
