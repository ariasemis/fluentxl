using System.Text.RegularExpressions;

namespace FluentXL.Demo.Extensions
{
    internal static class StringExtensions
    {
        /// <summary>
        /// Returns a new string with only alpha numeric characters
        /// </summary>
        /// <param name="input"></param>
        /// <returns></returns>
        public static string ToAlphanumeric(this string input)
            => Regex.Replace(input, "[^a-zA-Z0-9]", "");

        /// <summary>
        /// Split the string on spaces but preserving quoted text intact
        /// </summary>
        /// <param name="input"></param>
        /// <returns></returns>
        public static string[] SplitWithQuotedText(this string input)
            => Regex.Split(input, "(?<=^[^\"]*(?:\"[^\"]*\"[^\"]*)*) (?=(?:[^\"]*\"[^\"]*\")*[^\"]*$)");
    }
}
