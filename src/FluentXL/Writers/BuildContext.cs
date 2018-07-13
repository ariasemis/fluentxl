using FluentXL.Specifications;
using System;

namespace FluentXL.Writers
{
    internal class BuildContext : IBuildContext
    {
        public IStylesheetBuilder Stylesheet => throw new NotImplementedException();
    }
}
