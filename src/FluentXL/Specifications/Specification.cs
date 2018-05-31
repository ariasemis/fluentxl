namespace FluentXL.Specifications
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
        public static Sheets.IExpectSheetName Sheet()
            => Sheets.SheetSpecification.New();

        /// <summary>
        /// Creates a new column specification
        /// </summary>
        /// <returns>an instance of IColumnSpecification</returns>
        public static Columns.IColumnSpecification Column()
            => Columns.ColumnSpecification.New();

        /// <summary>
        /// Creates a new row specification
        /// </summary>
        /// <returns>an instance of IExpectRowIndex</returns>
        public static Rows.IExpectRowIndex Row()
            => Rows.RowSpecification.New();

        /// <summary>
        /// Creates a new cell specification
        /// </summary>
        /// <returns>an instance of IExpectCellColumn</returns>
        public static Cells.IExpectCellColumn Cell()
            => Cells.CellSpecification.New();

        /// <summary>
        /// Creates a new merge cell specification
        /// </summary>
        /// <returns>an instance of IExpectFrom</returns>
        public static MergeCells.IExpectFrom MergeCell()
            => MergeCells.MergeCellSpecification.New();
    }
}
