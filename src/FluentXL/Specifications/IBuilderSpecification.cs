namespace FluentXL.Specifications
{
    public interface IBuilderSpecification<out T>
    {
        T Build();
    }
}
