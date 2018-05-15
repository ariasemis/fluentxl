using FluentXL.Models;
using System.Collections.Generic;
using System.Linq;

namespace FluentXL.Specifications.Rows
{
    public class RowSpecification : IRowSpecification
    {
        private uint Index { get; set; }
        private IEnumerable<IBuilderSpecification<Cell>> CellSpecifications { get; set; }

        private RowSpecification() { }

        public static IRowSpecification Row()
            => new RowSpecification { CellSpecifications = Enumerable.Empty<IBuilderSpecification<Cell>>() };

        public IRowSpecification WithIndex(uint index)
        {
            return new RowSpecification
            {
                Index = index,
                CellSpecifications = CellSpecifications
            };
        }

        public IRowSpecification WithCell(IBuilderSpecification<Cell> cellSpecification)
        {
            return new RowSpecification
            {
                Index = Index,
                CellSpecifications = CellSpecifications.Append(cellSpecification)
            };
        }

        public IRowSpecification WithCells(IEnumerable<IBuilderSpecification<Cell>> specifications)
        {
            return new RowSpecification
            {
                Index = Index,
                CellSpecifications = CellSpecifications.Concat(specifications)
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
