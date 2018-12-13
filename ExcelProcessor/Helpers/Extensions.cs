using System;
using System.Collections.Generic;
using System.Text;
using System.Text.RegularExpressions;

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

        public static string Substring2(this string source, string left, string right)
        {
            return Regex.Match(source,string.Format("{0}(.*){1}", left, right)).Groups[1].Value;
        }

        public static bool IsNullOrEmpty(this string source)
        {
            return string.IsNullOrEmpty(source);
        }
    }
}
