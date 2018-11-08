using System;
using System.Collections.Generic;
using System.Text;

namespace ExcelProcessor.Helpers
{
    public static class Extensions
    {
        public static  bool Compare(this string str, string b)
        {
            return string.Equals(str, b, StringComparison.CurrentCultureIgnoreCase);
        }
    }
}
