using FluentXL.Models;
using System;
using System.Collections.Generic;
using System.Linq;

namespace FluentXL.Specifications.Sheets
{
    public class SheetSpecification : ISheetSpecification
    {
        private string name;
        private IEnumerable<IBuilderSpecification<Column>> columnSpecifications;
        private IEnumerable<IBuilderSpecification<Row>> rowSpecifications;
        private IEnumerable<IBuilderSpecification<MergeCell>> mergeCellSpecifications;

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

        public static ISheetSpecification New()
            => new SheetSpecification(
                string.Empty,
                Enumerable.Empty<IBuilderSpecification<Column>>(),
                Enumerable.Empty<IBuilderSpecification<Row>>(),
                Enumerable.Empty<IBuilderSpecification<MergeCell>>());

        public ISheetSpecification WithName(string name)
            => new SheetSpecification(name, columnSpecifications, rowSpecifications, mergeCellSpecifications);

        public ISheetSpecification WithColumn(IBuilderSpecification<Column> columnSpecification)
        {
            var cols = columnSpecifications.Append(columnSpecification);
            return new SheetSpecification(name, cols, rowSpecifications, mergeCellSpecifications);
        }

        public ISheetSpecification WithColumns<T>(IEnumerable<T> source, Func<T, uint, IBuilderSpecification<Column>> toSpecification)
        {
            // we add 1 to the index because it must start with 1
            var cols = columnSpecifications
                .Concat(source
                .Select((item, index) => toSpecification(item, (uint)index + 1)));
            return new SheetSpecification(name, cols, rowSpecifications, mergeCellSpecifications);
        }

        public ISheetSpecification WithColumns(IEnumerable<IBuilderSpecification<Column>> specifications)
        {
            var cols = columnSpecifications.Concat(specifications);
            return new SheetSpecification(name, cols, rowSpecifications, mergeCellSpecifications);
        }

        public ISheetSpecification WithRow(IBuilderSpecification<Row> rowSpecification)
        {
            var rows = rowSpecifications.Append(rowSpecification);
            return new SheetSpecification(name, columnSpecifications, rows, mergeCellSpecifications);
        }

        public ISheetSpecification WithRows<T>(IEnumerable<T> source, Func<T, uint, IBuilderSpecification<Row>> toSpecification)
        {
            // we add 1 to the index because it must start with 1
            var rows = rowSpecifications
                .Concat(source
                .Select((item, index) => toSpecification(item, (uint)index + 1)));
            return new SheetSpecification(name, columnSpecifications, rows, mergeCellSpecifications);
        }

        public ISheetSpecification WithRows(IEnumerable<IBuilderSpecification<Row>> specifications)
        {
            var rows = rowSpecifications.Concat(specifications);
            return new SheetSpecification(name, columnSpecifications, rows, mergeCellSpecifications);
        }

        public ISheetSpecification WithMergedCell(IBuilderSpecification<MergeCell> specification)
        {
            var merge = mergeCellSpecifications.Append(specification);
            return new SheetSpecification(name, columnSpecifications, rowSpecifications, merge);
        }

        public ISheetSpecification WithMergedCells(IEnumerable<IBuilderSpecification<MergeCell>> specifications)
        {
            var merge = mergeCellSpecifications.Concat(specifications);
            return new SheetSpecification(name, columnSpecifications, rowSpecifications, merge);
        }

        public Sheet Build()
        {
            var mergeCells = mergeCellSpecifications.Select(x => x.Build()).ToList();

            return new Sheet(
                name,
                columnSpecifications.Select(x => x.Build()),
                rowSpecifications.Select(x => x.Build()),
                new MergeCellCollection(mergeCells.Count, mergeCells));
        }
    }
}
