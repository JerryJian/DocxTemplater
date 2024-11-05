#if NET40_OR_GREATER
using System;

namespace DocxTemplater
{
    internal static class InternalExtensions
    {
        public static void ThrowIfNull(object argument, string paramName)
        {
            if (argument is null)
            {
                throw new ArgumentNullException(paramName);
            }
        }
    }
}
#endif