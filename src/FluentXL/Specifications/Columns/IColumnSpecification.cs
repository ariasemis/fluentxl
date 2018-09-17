using FluentXL.Elements;

namespace FluentXL.Specifications.Columns
{
    public interface IColumnSpecification
    {
        /// <summary>
        /// Specifies a column width.
        /// </summary>
        /// <param name="index">The index of the column.</param>
        /// <param name="width">The width of the column.</param>
        /// <returns>A new instance of IBuilderSpecification<Column>.</returns>
        IBuilderSpecification<Column> With(uint index, double width);
    }
}
