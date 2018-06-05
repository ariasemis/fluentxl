namespace FluentXL
{
    /// <summary>
    /// This class works as a common access point to the different specifications used by the fluentXL library.
    /// </summary>
    public static class Specification
    {
        /// <summary>
        /// Creates a new sheet specification
        /// </summary>
        /// <returns>an instance of IExpectSheetName</returns>
        public static Specifications.Sheets.IExpectSheetName Sheet()
            => Specifications.Sheets.SheetSpecification.New();

        /// <summary>
        /// Creates a new column specification
        /// </summary>
        /// <returns>an instance of IColumnSpecification</returns>
        public static Specifications.Columns.IColumnSpecification Column()
            => Specifications.Columns.ColumnSpecification.New();

        /// <summary>
        /// Creates a new row specification
        /// </summary>
        /// <returns>an instance of IExpectRowIndex</returns>
        public static Specifications.Rows.IExpectRowIndex Row()
            => Specifications.Rows.RowSpecification.New();

        /// <summary>
        /// Creates a new cell specification
        /// </summary>
        /// <returns>an instance of IExpectCellColumn</returns>
        public static Specifications.Cells.IExpectCellColumn Cell()
            => Specifications.Cells.CellSpecification.New();

        /// <summary>
        /// Creates a new merge cell specification
        /// </summary>
        /// <returns>an instance of IExpectFrom</returns>
        public static Specifications.MergeCells.IExpectFrom MergeCell()
            => Specifications.MergeCells.MergeCellSpecification.New();
    }
}
