using FluentXL.Models;
using System.Collections.Generic;

namespace FluentXL.Specifications.Rows
{
    public interface IExpectRowIndex
    {
        IExpectCells OnIndex(uint index);
    }

    public interface IExpectCells : IBuilderSpecification<Row>
    {
        IExpectCells WithCell(IBuilderSpecification<CellDefinition> cellSpecification);

        IExpectCells WithCells(IEnumerable<IBuilderSpecification<CellDefinition>> specifications);
    }
}
