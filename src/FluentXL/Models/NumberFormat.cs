﻿using System;

namespace FluentXL.Models
{
    public sealed class NumberFormat : IEquatable<NumberFormat>
    {
        public const uint INITIAL_ID = 164;

        public NumberFormat(
            uint id,
            string formatCode)
        {
            Id = id;
            FormatCode = formatCode ?? throw new ArgumentNullException(nameof(formatCode));
        }

        public uint Id { get; }

        public string FormatCode { get; }

        public static bool operator ==(NumberFormat a, NumberFormat b)
        {
            if (ReferenceEquals(a, b))
                return true;

            if (a is null || b is null)
                return false;

            return Equals(a, b);
        }

        public static bool operator !=(NumberFormat a, NumberFormat b)
            => !(a == b);

        public bool Equals(NumberFormat other)
        {
            if (ReferenceEquals(this, other))
                return true;
            if (other is null)
                return false;

            return FormatCode == other.FormatCode;
        }

        public override bool Equals(object obj)
            => obj is NumberFormat && Equals(obj as NumberFormat);

        public override int GetHashCode()
            => FormatCode.GetHashCode();
    }
}
