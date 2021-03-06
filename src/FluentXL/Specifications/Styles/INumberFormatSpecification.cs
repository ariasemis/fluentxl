﻿using FluentXL.Elements;

namespace FluentXL.Specifications.Styles
{
    public interface INumberFormatSpecification
    {
        /// <summary>
        /// Specifies the numbering format.
        /// </summary>
        /// <param name="format">The format to be used</param>
        /// <returns>An instance of <see cref="IBuilderSpecification{NumberFormat}"/></returns>
        IBuilderSpecification<NumberFormat> WithFormat(string format);

        /// <summary>
        /// Specifies the numbering format.
        /// </summary>
        /// <param name="format">The format to be used</param>
        /// <returns>An instance of <see cref="IBuilderSpecification{NumberFormat}"/></returns>
        IBuilderSpecification<NumberFormat> WithFormat(StandardNumberFormat format);
    }
}
