using System.Collections;
using System.Collections.Generic;

namespace FluentXL.Utils
{
    public class CountedCollection<T> : IEnumerable<T>
    {
        public CountedCollection(int count, IEnumerable<T> items)
        {
            Count = count;
            Items = items;
        }

        public int Count { get; }
        public IEnumerable<T> Items { get; }

        public IEnumerator<T> GetEnumerator()
            => Items.GetEnumerator();

        IEnumerator IEnumerable.GetEnumerator()
            => GetEnumerator();
    }
}
