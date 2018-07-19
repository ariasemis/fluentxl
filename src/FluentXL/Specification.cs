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

        /// <summary>
        /// Creates a new alignment specification
        /// </summary>
        /// <returns>An instance of IAlignmentSpecification</returns>
        public static Specifications.Styles.IAlignmentSpecification Alignment()
            => Specifications.Styles.AlignmentSpecification.New();

        /// <summary>
        /// Creates a new border specification
        /// </summary>
        /// <returns>An instance of IExpectBorderOutlineSpecification</returns>
        public static Specifications.Styles.IExpectBorderOutlineSpecification Border()
            => Specifications.Styles.BorderSpecification.New();

        /// <summary>
        /// Creates a new cell format specification
        /// </summary>
        /// <returns>An instance of ICellFormatSpecification</returns>
        public static Specifications.Styles.ICellFormatSpecification CellFormat()
            => Specifications.Styles.CellFormatSpecification.New();

        /// <summary>
        /// Creates a new color specification
        /// </summary>
        /// <returns>An instance of IColorSpecification</returns>
        public static Specifications.Styles.IColorSpecification Color()
            => Specifications.Styles.ColorSpecification.New();

        /// <summary>
        /// Creates a new fill specification
        /// </summary>
        /// <returns>An instance of IExpectFillPattern</returns>
        public static Specifications.Styles.IExpectFillPattern Fill()
            => Specifications.Styles.FillSpecification.New();

        /// <summary>
        /// Creates a new font specification
        /// </summary>
        /// <returns>An instance of IExpectFontName</returns>
        public static Specifications.Styles.IExpectFontName Font()
            => Specifications.Styles.FontSpecification.New();

        /// <summary>
        /// Creates a new number format specification
        /// </summary>
        /// <returns>An instance of INumberFormatSpecification</returns>
        public static Specifications.Styles.INumberFormatSpecification NumberFormat()
            => Specifications.Styles.NumberFormatSpecification.New();
    }
}
