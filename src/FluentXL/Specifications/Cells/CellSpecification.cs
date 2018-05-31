using FluentXL.Models;
using System;
using System.Globalization;

namespace FluentXL.Specifications.Cells
{
    public class CellSpecification : IExpectCellColumn, IExpectCellContent, IBuilderSpecification<CellDefinition>
    {
        private uint Column { get; set; }
        private CellType CellType { get; set; }
        private string Content { get; set; }

        private CellSpecification() { }

        public static IExpectCellColumn New()
            => new CellSpecification();

        public IExpectCellContent OnColumn(uint index)
            => new CellSpecification { Column = index };

        public IBuilderSpecification<CellDefinition> WithContent(DateTime value)
            => WithContent(value.ToOADate().ToString(CultureInfo.InvariantCulture), CellType.Date);

        public IBuilderSpecification<CellDefinition> WithContent(int value)
            => WithContent(value.ToString(CultureInfo.InvariantCulture), CellType.Number);

        public IBuilderSpecification<CellDefinition> WithContent(decimal value)
        {
            if (value == 0) value = Math.Truncate(value);
            return WithContent(value.ToString(CultureInfo.InvariantCulture), CellType.Number);
        }

        public IBuilderSpecification<CellDefinition> WithContent(bool value)
            => WithContent(value ? "1" : "0", CellType.Boolean);

        public IBuilderSpecification<CellDefinition> WithContent(string value)
            => WithContent(value, CellType.String);

        public IBuilderSpecification<CellDefinition> WithContent(string value, CellType type)
        {
            if (value == null)
                throw new ArgumentNullException(nameof(value));

            return new CellSpecification { Column = Column, CellType = type, Content = value };
        }

        public CellDefinition Build()
        {
            return new CellDefinition(
                Column,
                CellType,
                Content);
        }
    }
}
