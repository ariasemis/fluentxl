using FluentXL.Writers;

namespace FluentXL
{
    /// <summary>
    /// Utility class that provides methods to work with Spreadsheet documents.
    /// </summary>
    public static class Spreadsheet
    {
        /// <summary>
        /// Creates a new instance of the ISpreadsheetWriter interface.
        /// </summary>
        /// <returns>A new instance of ISpreadsheetWriter.</returns>
        public static ISpreadsheetWriter New()
            => new SpreadsheetWriter();
    }
}
