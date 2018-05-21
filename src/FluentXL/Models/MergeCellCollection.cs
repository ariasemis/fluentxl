using System.Collections;
using System.Collections.Generic;

namespace FluentXL.Models
{
    public class MergeCellCollection : IEnumerable<MergeCell>
    {
        public MergeCellCollection(int count, IEnumerable<MergeCell> mergeCells)
        {
            Count = count;
            MergeCells = mergeCells;
        }

        public int Count { get; }
        public IEnumerable<MergeCell> MergeCells { get; }

        public IEnumerator<MergeCell> GetEnumerator()
            => MergeCells.GetEnumerator();

        IEnumerator IEnumerable.GetEnumerator()
            => GetEnumerator();
    }
}
