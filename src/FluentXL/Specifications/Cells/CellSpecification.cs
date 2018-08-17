using FluentXL.Models;
using System;
using System.Globalization;

namespace FluentXL.Specifications.Cells
{
    public class CellSpecification : IExpectCellColumn, IExpectCellContent, IExpectCellFormat, IExpectCellRow, IBuilderSpecification<Cell>
    {
        private uint Row { get; set; }
        private uint Column { get; set; }
        private CellType CellType { get; set; }
        private string Content { get; set; }
        private IBuilderSpecification<CellFormat> CellFormat { get; set; } = null;

        private CellSpecification() { }

        public static IExpectCellColumn New()
            => new CellSpecification();

        public IExpectCellContent OnColumn(uint index)
            => new CellSpecification { Column = index };

        public IExpectCellFormat WithContent(DateTime value)
            //TODO: for datetime, add a default style with short date numbering format
            => WithContent(value.ToOADate().ToString(CultureInfo.InvariantCulture), CellType.Date);

        public IExpectCellFormat WithContent(int value)
            => WithContent(value.ToString(CultureInfo.InvariantCulture), CellType.Number);

        public IExpectCellFormat WithContent(decimal value)
        {
            if (value == 0) value = Math.Truncate(value);
            return WithContent(value.ToString(CultureInfo.InvariantCulture), CellType.Number);
        }

        public IExpectCellFormat WithContent(bool value)
            => WithContent(value ? "1" : "0", CellType.Boolean);

        public IExpectCellFormat WithContent(string value)
            => WithContent(value, CellType.String);

        public IExpectCellFormat WithContent(string value, CellType type)
        {
            if (value == null)
                throw new ArgumentNullException(nameof(value));

            return new CellSpecification { Column = Column, CellType = type, Content = value };
        }

        public IExpectCellRow WithStyle(IBuilderSpecification<CellFormat> formatSpecification)
        {
            return new CellSpecification { Column = Column, CellType = CellType, Content = Content, CellFormat = formatSpecification };
        }

        public IBuilderSpecification<Cell> OnRow(uint index)
        {
            return new CellSpecification { Column = Column, CellType = CellType, Content = Content, CellFormat = CellFormat, Row = index };
        }

        public Cell Build(IBuildContext context)
        {
            var cellFormat = CellFormat?.Build(context);

            return new Cell(
                Row,
                Column,
                CellType,
                Content,
                cellFormat == null ? (uint?)null : context.Stylesheet.AddCellFormat(cellFormat));
        }
    }
}
