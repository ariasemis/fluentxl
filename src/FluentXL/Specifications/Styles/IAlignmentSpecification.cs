using FluentXL.Models;

namespace FluentXL.Specifications.Styles
{
    public interface IAlignmentSpecification : IBuilderSpecification<Alignment>
    {
        /// <summary>
        /// Specifies the horizontal alignment.
        /// </summary>
        /// <param name="horizontalAlignment">Horizontal alignment</param>
        /// <returns>A new instance of IAlignmentSpecification</returns>
        IAlignmentSpecification WithHorizontal(HorizontalAlignment horizontalAlignment);

        /// <summary>
        /// Specifies the vertical alignment.
        /// </summary>
        /// <param name="verticalAlignment">Vertical alignment</param>
        /// <returns>A new instance of IAlignmentSpecification</returns>
        IAlignmentSpecification WithVertical(VerticalAlignment verticalAlignment);

        /// <summary>
        /// Specifies the reading order of the alignment.
        /// </summary>
        /// <param name="readingOrder">The reading order</param>
        /// <returns>A new instance of IAlignmentSpecification</returns>
        IAlignmentSpecification WithReadingOrder(ReadingOrder readingOrder);

        /// <summary>
        /// Specifies that the container must shrink to fit the content.
        /// </summary>
        /// <returns>A new instance of IAlignmentSpecification</returns>
        IAlignmentSpecification ShrinkToFit();

        /// <summary>
        /// Specifies that the text must be wrapped.
        /// </summary>
        /// <returns>A new instance of IAlignmentSpecification</returns>
        IAlignmentSpecification WrapText();
    }
}
