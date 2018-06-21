using System;
using System.Collections.Generic;
using System.Text;

namespace FluentXL.Specifications.CellFormats
{
    public interface IExpectFontName
    {
        IExpectFontSize WithFont(string name);
    }

    public interface IExpectFontSize : IExpectFontColor
    {
        IExpectFontColor WithSize(double size);
    }

    public interface IExpectFontColor : IExpectFontOptions
    {
        IExpectFontOptions WithColor(IColorSpecification colorSpecification);
    }

    public interface IExpectFontOptions
    {
        IExpectFontOptions Bold();

        IExpectFontOptions Italic();
    }
}
