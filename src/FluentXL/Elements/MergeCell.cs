using System;

namespace FluentXL.Elements
{
    public class MergeCell
    {
        public MergeCell(string reference)
        {
            if (string.IsNullOrEmpty(reference))
                throw new ArgumentNullException(nameof(reference));

            Reference = reference;
        }

        public string Reference { get; }
    }
}
