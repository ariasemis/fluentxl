using FluentXL.Demo.Samples;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Reflection;

namespace FluentXL.Demo.Extensions
{
    public static class AssemblyExtensions
    {
        public static ISample ResolveSample(this Assembly assembly, string name)
        {
            if (assembly == null)
                throw new ArgumentNullException(nameof(assembly));
            if (string.IsNullOrEmpty(name))
                throw new ArgumentNullException(nameof(name));

            var typeName = name.ToAlphanumeric().ToLowerInvariant();

            var contract = typeof(ISample);
            var type = assembly
                .GetTypes()
                .Where(x => contract.IsAssignableFrom(x) && x.Name.Equals(typeName, StringComparison.InvariantCultureIgnoreCase))
                .FirstOrDefault();

            return type == null ? null : Activator.CreateInstance(type) as ISample;
        }

        public static IEnumerable<string> GetSampleNames(this Assembly assembly)
        {
            var contract = typeof(ISample);
            return assembly
                .GetTypes()
                .Where(x => contract.IsAssignableFrom(x) && x.IsClass)
                .Select(x => x.Name.ToKebabCase());
        }
    }
}
