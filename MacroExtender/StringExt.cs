using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;

namespace MacroExtender
{
    public static class StringExt
    {
        public static string Left(this string value, int maxLength)
        {
            return value.Length <= maxLength ? value : value.Substring(0, maxLength);
        }

        public static string Right(this string value, int maxLength)
        {
            return value.Length <= maxLength ? value : value.Substring(value.Length - maxLength, maxLength);
        }
    }
}
