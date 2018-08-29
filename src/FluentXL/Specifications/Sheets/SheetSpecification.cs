using FluentXL.Models;
using FluentXL.Utils;
using System;
using System.Collections.Generic;
using System.Linq;

namespace FluentXL.Specifications.Sheets
{
    public class SheetSpecification : IExpectSheetName, IExpectSheetContent
    {
        private readonly string name;
        private readonly IEnumerable<IBuilderSpecification<Column>> columnSpecifications;
        private readonly IEnumerable<IBuilderSpecification<Row>> rowSpecifications;
        private readonly IEnumerable<IBuilderSpecification<MergeCell>> mergeCellSpecifications;

        private SheetSpecification(
            string name,
            IEnumerable<IBuilderSpecification<Column>> columnSpecifications,
            IEnumerable<IBuilderSpecification<Row>> rowSpecifications,
            IEnumerable<IBuilderSpecification<MergeCell>> mergeCellSpecifications)
        {
            this.name = name;
            this.columnSpecifications = columnSpecifications;
            this.rowSpecifications = rowSpecifications;
            this.mergeCellSpecifications = mergeCellSpecifications;
        }

        public static IExpectSheetName New()
            => new SheetSpecification(
                string.Empty,
                Enumerable.Empty<IBuilderSpecification<Column>>(),
                Enumerable.Empty<IBuilderSpecification<Row>>(),
                Enumerable.Empty<IBuilderSpecification<MergeCell>>());

        public IExpectSheetContent WithName(string name)
        {
            if (name == null) throw new ArgumentNullException(nameof(name));

            return new SheetSpecification(name, columnSpecifications, rowSpecifications, mergeCellSpecifications);
        }

        public IExpectSheetContent WithColumn(IBuilderSpecification<Column> columnSpecification)
        {
            var cols = columnSpecifications.Append(columnSpecification);
            return new SheetSpecification(name, cols, rowSpecifications, mergeCellSpecifications);
        }

        public IExpectSheetContent WithColumns<T>(IEnumerable<T> source, Func<T, uint, IBuilderSpecification<Column>> toSpecification)
        {
            // we add 1 to the index because it must start with 1
            var cols = columnSpecifications
                .Concat(source
                .Select((item, index) => toSpecification(item, (uint)index + 1)));
            return new SheetSpecification(name, cols, rowSpecifications, mergeCellSpecifications);
        }

        public IExpectSheetContent WithColumns(IEnumerable<IBuilderSpecification<Column>> specifications)
        {
            var cols = columnSpecifications.Concat(specifications);
            return new SheetSpecification(name, cols, rowSpecifications, mergeCellSpecifications);
        }

        public IExpectSheetContent WithRow(IBuilderSpecification<Row> rowSpecification)
        {
            var rows = rowSpecifications.Append(rowSpecification);
            return new SheetSpecification(name, columnSpecifications, rows, mergeCellSpecifications);
        }

        public IExpectSheetContent WithRows<T>(IEnumerable<T> source, Func<T, uint, IBuilderSpecification<Row>> toSpecification)
        {
            // we add 1 to the index because it must start with 1
            var rows = rowSpecifications
                .Concat(source
                .Select((item, index) => toSpecification(item, (uint)index + 1)));
            return new SheetSpecification(name, columnSpecifications, rows, mergeCellSpecifications);
        }

        public IExpectSheetContent WithRows(IEnumerable<IBuilderSpecification<Row>> specifications)
        {
            var rows = rowSpecifications.Concat(specifications);
            return new SheetSpecification(name, columnSpecifications, rows, mergeCellSpecifications);
        }

        public IExpectSheetContent WithMergedCell(IBuilderSpecification<MergeCell> specification)
        {
            var merge = mergeCellSpecifications.Append(specification);
            return new SheetSpecification(name, columnSpecifications, rowSpecifications, merge);
        }

        public IExpectSheetContent WithMergedCells(IEnumerable<IBuilderSpecification<MergeCell>> specifications)
        {
            var merge = mergeCellSpecifications.Concat(specifications);
            return new SheetSpecification(name, columnSpecifications, rowSpecifications, merge);
        }

        public Sheet Build(IBuildContext context)
        {
            var mergeCells = mergeCellSpecifications.Select(x => x.Build(context)).ToList();

            return new Sheet(
                name,
                columnSpecifications.Select(x => x.Build(context)),
                rowSpecifications.Select(x => x.Build(context)),
                new CountedCollection<MergeCell>(mergeCells.Count, mergeCells));
        }
    }
}
