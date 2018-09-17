using FluentXL.Elements;
using System;

namespace FluentXL.Specifications.Columns
{
    public class ColumnSpecification :
        IColumnSpecification,
        IBuilderSpecification<Column>
    {
        private uint Index { get; set; }
        private double Width { get; set; }

        private ColumnSpecification() { }

        public static IColumnSpecification New()
            => new ColumnSpecification();

        public IBuilderSpecification<Column> With(uint index, double width)
        {
            if (index < Column.MIN_INDEX)
                throw new ArgumentException($"The index value must be greater or equal to {Column.MIN_INDEX}", nameof(index));
            if (index > Column.MAX_INDEX)
                throw new ArgumentException($"The index value must be less or equal to {Column.MAX_INDEX}", nameof(index));

            return new ColumnSpecification
            {
                Index = index,
                Width = width
            };
        }

        public Column Build(IBuildContext context)
            => new Column(Index, Width);
    }
}
