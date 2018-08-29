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
            => new CellSpecification(this) { Column = index };

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
            => new CellSpecification(this) { Row = index };

        public Cell Build(IBuildContext context)
        {
            var cellFormat = Compose(DefaultCellFormat, CellFormat)(context);

            return new Cell(
                Row,
                Column,
                CellType,
                Content,
                cellFormat == null ? (uint?)null : context.Stylesheet.Add(cellFormat));
        }

        private static Func<IBuildContext, CellFormat> Compose(IBuilderSpecification<CellFormat> defaultSpec, IBuilderSpecification<CellFormat> spec)
        {
            return (context) =>
            {
                var defaultFormat = defaultSpec?.Build(context);
                var format = spec?.Build(context);

                if (format == null && defaultFormat == null)
                    return null;

                return new CellFormat(
                    formatId: format?.FormatId ?? defaultFormat?.FormatId,
                    fontId: format?.FontId ?? defaultFormat?.FontId,
                    fillId: format?.FillId ?? defaultFormat?.FillId,
                    borderId: format?.BorderId ?? defaultFormat?.BorderId,
                    numberFormatId: format?.NumberFormatId ?? defaultFormat?.NumberFormatId,
                    hasPivotButton: format?.HasPivotButton ?? defaultFormat?.HasPivotButton,
                    hasQuotePrefix: format?.HasQuotePrefix ?? defaultFormat?.HasQuotePrefix,
                    alignment: format?.Alignment ?? defaultFormat?.Alignment,
                    protection: format?.Protection ?? defaultFormat?.Protection);
            };
        }
    }
}
