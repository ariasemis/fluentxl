using System.Text.RegularExpressions;

namespace FluentXL.Demo.Extensions
{
    internal static class StringExtensions
    {
        public static string ToAlphanumeric(this string input)
            => Regex.Replace(input, "[^a-zA-Z0-9]", "");
    }
}
