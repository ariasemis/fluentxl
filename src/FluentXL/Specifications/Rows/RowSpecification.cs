using FluentXL.Models;
using FluentXL.Specifications.Cells;
using System.Collections.Generic;
using System.Linq;

namespace FluentXL.Specifications.Rows
{
    public class RowSpecification : IExpectRowIndex, IExpectCells
    {
        private uint Index { get; set; }
        private IEnumerable<IBuilderSpecification<Cell>> CellSpecifications { get; set; }

        private RowSpecification() { }

        public static IExpectRowIndex New()
            => new RowSpecification { CellSpecifications = Enumerable.Empty<IBuilderSpecification<Cell>>() };

        public IExpectCells OnIndex(uint index)
        {
            return new RowSpecification
            {
                Index = index,
                CellSpecifications = CellSpecifications
            };
        }

        public IExpectCells WithCell(IExpectCellRow cellSpecification)
        {
            return new RowSpecification
            {
                Index = Index,
                CellSpecifications = CellSpecifications.Append(cellSpecification.OnRow(Index))
            };
        }

        public IExpectCells WithCells(IEnumerable<IExpectCellRow> specifications)
        {
            return new RowSpecification
            {
                Index = Index,
                CellSpecifications = CellSpecifications.Concat(specifications.Select(x => x.OnRow(Index)))
            };
        }

        public Row Build()
        {
            return new Row(
                Index,
                CellSpecifications.Select(x => x.Build()));
        }
    }
}
