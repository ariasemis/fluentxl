﻿using FluentXL.Models;

namespace FluentXL.Specifications.Columns
{
    public class ColumnSpecification :
        IColumnSpecification,
        IBuilderSpecification<Column>
    {
        private uint Index { get; set; }
        private double Width { get; set; }

        private ColumnSpecification() { }

        public static IColumnSpecification Column()
            => new ColumnSpecification();

        public IBuilderSpecification<Column> With(uint index, double width)
        {
            return new ColumnSpecification
            {
                Index = index,
                Width = width
            };
        }

        public Column Build()
        {
            return new Column(Index, Width);
        }
    }
}
