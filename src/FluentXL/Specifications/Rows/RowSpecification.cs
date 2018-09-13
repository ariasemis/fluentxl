using FluentXL.Models;
using FluentXL.Specifications.Cells;
using System;
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
            if (index < Row.MIN_INDEX)
                throw new ArgumentException($"The index value must be greater or equal to {Row.MIN_INDEX}", nameof(index));
            if (index > Row.MAX_INDEX)
                throw new ArgumentException($"The index value must be less or equal to {Row.MAX_INDEX}", nameof(index));

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

        public Row Build(IBuildContext context)
        {
            return new Row(
                Index,
                CellSpecifications.Select(x => x.Build(context)));
        }
    }
}
