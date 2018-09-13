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
        private IBuilderSpecification<CellFormat> DefaultCellFormat { get; set; } = null;

        private CellSpecification() { }

        private CellSpecification(CellSpecification cellSpecification)
        {
            Row = cellSpecification.Row;
            Column = cellSpecification.Column;
            CellType = cellSpecification.CellType;
            Content = cellSpecification.Content;
            CellFormat = cellSpecification.CellFormat;
            DefaultCellFormat = cellSpecification.DefaultCellFormat;
        }

        public static IExpectCellColumn New()
            => new CellSpecification();

        public IExpectCellContent OnColumn(uint index)
        {
            if (index < Models.Column.MIN_INDEX)
                throw new ArgumentException($"The index value must be greater or equal to {Models.Column.MIN_INDEX}", nameof(index));
            if (index > Models.Column.MAX_INDEX)
                throw new ArgumentException($"The index value must be less or equal to {Models.Column.MAX_INDEX}", nameof(index));

            return new CellSpecification(this) { Column = index };
        }

        public IExpectCellFormat WithContent(DateTime value)
        {
            DefaultCellFormat = Specification.CellFormat()
                .WithNumberFormat(Specification.NumberFormat().WithFormat(StandardNumberFormat.ShortDate));

            return WithContent(value.ToOADate().ToString(CultureInfo.InvariantCulture), CellType.Date);
        }

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

            return new CellSpecification(this) { CellType = type, Content = value };
        }

        public IExpectCellRow WithStyle(IBuilderSpecification<CellFormat> formatSpecification)
            => new CellSpecification(this) { CellFormat = formatSpecification };

        public IBuilderSpecification<Cell> OnRow(uint index)
        {
            if (index < Models.Row.MIN_INDEX)
                throw new ArgumentException($"The index value must be greater or equal to {Models.Row.MIN_INDEX}", nameof(index));
            if (index > Models.Row.MAX_INDEX)
                throw new ArgumentException($"The index value must be less or equal to {Models.Row.MAX_INDEX}", nameof(index));

            return new CellSpecification(this) { Row = index };
        }

        public Cell Build(IBuildContext context)
        {
            var formatSpec = new Styles.CompositeCellFormatSpecification(DefaultCellFormat, CellFormat);
            var cellFormat = formatSpec.Build(context);

            return new Cell(
                Row,
                Column,
                CellType,
                Content,
                cellFormat == null ? (uint?)null : context.Stylesheet.Add(cellFormat));
        }
    }
}
