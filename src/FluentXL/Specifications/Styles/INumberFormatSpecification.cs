using FluentXL.Models;

namespace FluentXL.Specifications.Styles
{
    public interface INumberFormatSpecification
    {
        IBuilderSpecification<NumberFormat> WithFormat(string format);

        IBuilderSpecification<NumberFormat> WithFormat(StandardNumberFormat format);
    }
}
