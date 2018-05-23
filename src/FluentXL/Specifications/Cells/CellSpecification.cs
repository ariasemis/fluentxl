using FluentXL.Models;
using System;
using System.Globalization;

namespace FluentXL.Specifications.Cells
{
    public class CellSpecification : IExpectCellColumn, IExpectCellContent, IBuilderSpecification<Cell>
    {
        private uint Row { get; set; }
        private uint Column { get; set; }
        private CellType CellType { get; set; }
        private string Content { get; set; }

        private CellSpecification() { }

        // TODO: find a way to get the row index from the parent so we dont need to specify it again
        public static IExpectCellColumn Cell(uint row)
            => new CellSpecification { Row = row };

        public IExpectCellContent WithColumn(uint index)
            => new CellSpecification { Row = Row, Column = index };

        public IBuilderSpecification<Cell> WithContent(DateTime value)
            => WithContent(value.ToOADate().ToString(CultureInfo.InvariantCulture), CellType.Date);

        public IBuilderSpecification<Cell> WithContent(int value)
            => WithContent(value.ToString(CultureInfo.InvariantCulture), CellType.Number);

        public IBuilderSpecification<Cell> WithContent(decimal value)
        {
            if (value == 0) value = Math.Truncate(value);
            return WithContent(value.ToString(CultureInfo.InvariantCulture), CellType.Number);
        }

        public IBuilderSpecification<Cell> WithContent(bool value)
            => WithContent(value ? "1" : "0", CellType.Boolean);

        public IBuilderSpecification<Cell> WithContent(string value)
            => WithContent(value, CellType.String);

        public IBuilderSpecification<Cell> WithContent(string value, CellType type)
            => new CellSpecification { Row = Row, Column = Column, CellType = type, Content = value };

        public Cell Build()
        {
            return new Cell(
                Row,
                Column,
                CellType,
                Content);
        }
    }
}
