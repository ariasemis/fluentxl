namespace FluentXL.Models
{
    public class Protection
    {
        public Protection(bool locked, bool hidden)
        {
            Locked = locked;
            Hidden = hidden;
        }

        public bool Locked { get; }

        public bool Hidden { get; set; }
    }
}
