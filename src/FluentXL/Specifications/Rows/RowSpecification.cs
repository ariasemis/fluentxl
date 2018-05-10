using FluentXL.Models;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;

namespace FluentXL.Specifications.Rows
{
    public class RowSpecification : IBuilderSpecification<Row>
    {
        private uint index;
        private IEnumerable<IBuilderSpecification<Cell>> cellSpecifications;

        public void WithIndex(uint index)
        {
            this.index = index;
        }

        public void WithCell(IBuilderSpecification<Cell> cellSpecification)
        {
            cellSpecifications = cellSpecifications.Append(cellSpecification);
        }

        public void WithCells(IEnumerable<IBuilderSpecification<Cell>> specifications)
        {
            cellSpecifications = cellSpecifications.Concat(specifications);
        }

        public Row Build()
        {
            throw new NotImplementedException();
        }
    }
}
