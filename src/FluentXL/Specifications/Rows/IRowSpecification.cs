using FluentXL.Models;
using System.Collections.Generic;

namespace FluentXL.Specifications.Rows
{
    public interface IRowSpecification : IBuilderSpecification<Row>
    {
        IRowSpecification WithIndex(uint index);

        IRowSpecification WithCell(IBuilderSpecification<CellDefinition> cellSpecification);

        IRowSpecification WithCells(IEnumerable<IBuilderSpecification<CellDefinition>> specifications);
    }
}
