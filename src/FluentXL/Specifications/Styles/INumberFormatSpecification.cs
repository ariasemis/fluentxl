using FluentXL.Models;

namespace FluentXL.Specifications.Styles
{
    public interface IExpectNumberFormatSpecification
    {
        IExpectNumberFormatStylesheetSpecification WithFormat(string format);

        IBuilderSpecification<NumberFormat> WithFormat(StandardNumberFormat format);
    }

    public interface IExpectNumberFormatStylesheetSpecification
    {
        IBuilderSpecification<NumberFormat> OnStylesheet(IStylesheetSpecification stylesheet);
    }
}
