using System;
using System.Collections.Generic;
using System.Text;

namespace ExcelProcessor.Helpers
{
    public static class Extensions
    {
        public static  bool CompareIgnoreCase(this string str, string b)
        {
            return string.Equals(str, b, StringComparison.OrdinalIgnoreCase);
        }

        public static string ReplaceSpace(this string str)
        {
            return str.Replace(" ", string.Empty);
        }        
    }
}
