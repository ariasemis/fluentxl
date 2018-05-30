using FluentXL.Models;
using System.Collections.Generic;
using System.Linq;

namespace FluentXL.Specifications.Rows
{
    public class RowSpecification : IRowSpecification
    {
        private uint Index { get; set; }
        private IEnumerable<IBuilderSpecification<CellDefinition>> CellSpecifications { get; set; }

        private RowSpecification() { }

        public static IRowSpecification Row()
            => new RowSpecification { CellSpecifications = Enumerable.Empty<IBuilderSpecification<CellDefinition>>() };

        public IRowSpecification WithIndex(uint index)
        {
            return new RowSpecification
            {
                Index = index,
                CellSpecifications = CellSpecifications
            };
        }

        public IRowSpecification WithCell(IBuilderSpecification<CellDefinition> cellSpecification)
        {
            return new RowSpecification
            {
                Index = Index,
                CellSpecifications = CellSpecifications.Append(cellSpecification)
            };
        }

        public IRowSpecification WithCells(IEnumerable<IBuilderSpecification<CellDefinition>> specifications)
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
                CellSpecifications.Select(x => new Cell(Index, x.Build())));
        }
    }
}
