namespace FluentXL.Writers
{
    internal interface IWriter<in T>
        where T : class
    {
        void Write(T element);
    }
}
