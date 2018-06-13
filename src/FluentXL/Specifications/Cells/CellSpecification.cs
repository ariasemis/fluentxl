using FluentXL.Models;
using System;
using System.Globalization;

namespace FluentXL.Specifications.Cells
{
    public class CellSpecification : IExpectCellColumn, IExpectCellContent, IExpectCellRow, IBuilderSpecification<Cell>
    {
        private uint Row { get; set; }
        private uint Column { get; set; }
        private CellType CellType { get; set; }
        private string Content { get; set; }

        private CellSpecification() { }

        public static IExpectCellColumn New()
            => new CellSpecification();

        public IExpectCellContent OnColumn(uint index)
            => new CellSpecification { Column = index };

        public IExpectCellRow WithContent(DateTime value)
            => WithContent(value.ToOADate().ToString(CultureInfo.InvariantCulture), CellType.Date);

        public IExpectCellRow WithContent(int value)
            => WithContent(value.ToString(CultureInfo.InvariantCulture), CellType.Number);

        public IExpectCellRow WithContent(decimal value)
        {
            if (value == 0) value = Math.Truncate(value);
            return WithContent(value.ToString(CultureInfo.InvariantCulture), CellType.Number);
        }

        public IExpectCellRow WithContent(bool value)
            => WithContent(value ? "1" : "0", CellType.Boolean);

        public IExpectCellRow WithContent(string value)
            => WithContent(value, CellType.String);

        public IExpectCellRow WithContent(string value, CellType type)
        {
            if (value == null)
                throw new ArgumentNullException(nameof(value));

            return new CellSpecification { Column = Column, CellType = type, Content = value };
        }

        public IBuilderSpecification<Cell> OnRow(uint index)
        {
            return new CellSpecification { Column = Column, CellType = CellType, Content = Content, Row = index };
        }

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
