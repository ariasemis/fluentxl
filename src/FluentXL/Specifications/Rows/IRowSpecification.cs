using FluentXL.Models;
using FluentXL.Specifications.Cells;
using System.Collections.Generic;

namespace FluentXL.Specifications.Rows
{
    public interface IExpectRowIndex
    {
        /// <summary>
        /// Specifies the index of the Row being built.
        /// </summary>
        /// <param name="index">The index value of the Row.</param>
        /// <returns>A new instance of IExpectCells.</returns>
        IExpectCells OnIndex(uint index);
    }

    public interface IExpectCells : IBuilderSpecification<Row>
    {
        /// <summary>
        /// Includes a cell specification to the Row being built.
        /// </summary>
        /// <param name="cellSpecification">A Cell specification.</param>
        /// <returns>A new instance of IExpectCells.</returns>
        IExpectCells WithCell(IExpectCellRow cellSpecification);

        /// <summary>
        /// Includes a series of cell specifications to the Row being built.
        /// </summary>
        /// <param name="specifications">A series of Cell specifications.</param>
        /// <returns>A new instance of IExpectCells.</returns>
        IExpectCells WithCells(IEnumerable<IExpectCellRow> specifications);
    }
}
