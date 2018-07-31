using System.Collections.Generic;

namespace FluentXL.Demo
{
    public class Input
    {
        public Input()
        {
            Arguments = new List<string>();
            Options = new Dictionary<string, string>();
            Flags = new List<string>();
        }

        public string Command { get; set; }
        public IList<string> Arguments { get; set; }
        public IDictionary<string, string> Options { get; set; }
        public IList<string> Flags { get; set; }
    }
}
