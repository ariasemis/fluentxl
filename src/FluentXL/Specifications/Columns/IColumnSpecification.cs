using FluentXL.Models;

namespace FluentXL.Specifications.Columns
{
    public interface IColumnSpecification
    {
        IBuilderSpecification<Column> With(uint index, double width);
    }
}
