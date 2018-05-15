using System.Collections.Generic;
using System.Linq;

namespace FluentXL.Utils
{
    internal static class ReferenceHelper
    {
        static readonly string[] alphabet = { string.Empty, "A", "B", "C", "D", "E", "F", "G", "H", "I", "J", "K", "L", "M", "N", "O", "P", "Q", "R", "S", "T", "U", "V", "W", "X", "Y", "Z" };

        static IEnumerable<string> excelStrings = CreateExcelStrings();

        static Dictionary<uint, string> columnsByIndex = excelStrings.Select((entry, index) => new { Entry = entry, Index = index + 1 })
                                                                    .ToDictionary(x => (uint)x.Index, x => x.Entry);

        static Dictionary<string, uint> indexByColumns = excelStrings.Select((entry, index) => new { Entry = entry, Index = index + 1 })
                                                                     .ToDictionary(x => x.Entry, x => (uint)x.Index);

        public static string GetReference(uint row, uint column)
        {
            return columnsByIndex[column] + row;
        }

        public static string GetNextReference(uint row, uint column)
        {
            return columnsByIndex[column++] + row;
        }

        public static string GetColumnLetter(uint index)
        {
            return columnsByIndex[index];
        }

        public static string GetColumnLetter(string reference)
        {
            string column = string.Empty;
            foreach (char c in reference)
            {
                if (!char.IsDigit(c))
                    column += c;
            }
            return column;
        }

        public static uint GetColumnIndex(string reference)
        {
            return indexByColumns[GetColumnLetter(reference)];
        }

        private static IEnumerable<string> CreateExcelStrings()
        {
            return from c1 in alphabet
                   from c2 in alphabet
                   from c3 in alphabet.Skip(1)
                   where string.IsNullOrEmpty(c1) || !string.IsNullOrEmpty(c2)
                   select c1 + c2 + c3;
        }
    }
}
