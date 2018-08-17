using System;

namespace FluentXL.Models
{
    public class Protection : IEquatable<Protection>
    {
        public Protection(bool locked, bool hidden)
        {
            Locked = locked;
            Hidden = hidden;
        }

        public bool Locked { get; }

        public bool Hidden { get; set; }

        public static bool operator ==(Protection a, Protection b)
        {
            if (ReferenceEquals(a, b))
                return true;

            if (a is null || b is null)
                return false;

            return Equals(a, b);
        }

        public static bool operator !=(Protection a, Protection b)
            => !(a == b);

        public bool Equals(Protection other)
        {
            if (ReferenceEquals(this, other))
                return true;
            if (other is null)
                return false;

            return Locked == other.Locked && Hidden == other.Hidden;
        }

        public override bool Equals(object obj)
            => obj is Protection && Equals(obj as Protection);

        public override int GetHashCode()
        {
            unchecked
            {
                int hash = 19;

                hash = hash * 29 + Locked.GetHashCode();
                hash = hash * 29 + Hidden.GetHashCode();

                return hash;
            }
        }
    }
}
