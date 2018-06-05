namespace FluentXL.Specifications.MergeCells
{
    public interface IExpectFrom
    {
        /// <summary>
        /// Specifies the initial cell of the range to be merged.
        /// </summary>
        /// <param name="row">The row index of the cell.</param>
        /// <param name="column">The column index of the cell.</param>
        /// <returns>A new instance of IExpectTo.</returns>
        IExpectTo From(uint row, uint column);
    }
}
