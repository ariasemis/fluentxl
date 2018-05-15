using FluentXL.Models;
using System.Collections.Generic;

namespace FluentXL.Specifications.Rows
{
    public interface IRowSpecification : IBuilderSpecification<Row>
    {
        IRowSpecification WithIndex(uint index);

        IRowSpecification WithCell(IBuilderSpecification<Cell> cellSpecification);

        IRowSpecification WithCells(IEnumerable<IBuilderSpecification<Cell>> specifications);
    }
}
