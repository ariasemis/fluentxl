using FluentXL.Specifications;
using FluentXL.Specifications.Styles;
using System;

namespace FluentXL.Writers
{
    internal class BuildContext : IBuildContext
    {
        public IStylesheetSpecification Stylesheet => throw new NotImplementedException();
    }
}
