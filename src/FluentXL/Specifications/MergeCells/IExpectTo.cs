using FluentXL.Elements;

namespace FluentXL.Specifications.MergeCells
{
    public interface IExpectTo
    {
        /// <summary>
        /// Specifies the final cell of the range to be merged.
        /// </summary>
        /// <param name="row">The row index of the cell.</param>
        /// <param name="column">The column index of the cell.</param>
        /// <returns>A new instance of <see cref="IBuilderSpecification{MergeCell}"/>.</returns>
        IBuilderSpecification<MergeCell> To(uint row, uint column);
    }
}
