using FluentXL.Elements;
using FluentXL.Specifications;
using System.IO;

namespace FluentXL
{
    /// <summary>
    /// Defines a class that allows the creation of new Spreadsheet documents.
    /// </summary>
    public interface ISpreadsheetWriter
    {
        /// <summary>
        /// Returns a new ISpreadsheetWriter that includes the sheet specification provided as an argument.
        /// </summary>
        /// <param name="specification">A Sheet specification to include in the resulting document.</param>
        /// <returns>A new instance of ISpreadsheetWriter.</returns>
        ISpreadsheetWriter WithSheet(IBuilderSpecification<Sheet> specification);

        /// <summary>
        /// Writes the Spreadsheet document to the stream provided as an argument.
        /// </summary>
        /// <param name="stream">The Stream the document should be written to.</param>
        void WriteTo(Stream stream);
    }
}
