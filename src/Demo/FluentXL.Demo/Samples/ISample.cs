namespace FluentXL.Demo.Samples
{
    public interface ISample
    {
        ISpreadsheetWriter Run();

        string GetInfo();
    }
}
