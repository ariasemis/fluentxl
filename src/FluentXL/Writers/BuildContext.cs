using FluentXL.Specifications;

namespace FluentXL.Writers
{
    internal class BuildContext : IBuildContext
    {
        public IStylesheetBuilder Stylesheet => new StylesheetBuilder();
    }
}
