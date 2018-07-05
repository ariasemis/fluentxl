namespace FluentXL.Specifications
{
    /// <summary>
    /// Represents a specification of an object to be built.
    /// </summary>
    /// <typeparam name="T">The type of the object specified.</typeparam>
    public interface IBuilderSpecification<out T>
    {
        /// <summary>
        /// Returns a new instance of the specified object.
        /// </summary>
        /// <param name="context">Context object with values needed to build the object</param>
        /// <returns>A new instance of type T.</returns>
        T Build(IBuildContext context);
    }
}
