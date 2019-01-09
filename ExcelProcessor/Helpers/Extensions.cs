using System;
using System.Collections.Generic;
using System.IO;
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

        public static string GetStringBetween(this string source, string left, string right)
        {
            return Regex.Match(source,string.Format("{0}(.*){1}", left, right)).Groups[1].Value;
        }

        public static bool IsNullOrEmpty(this string source)
        {
            return string.IsNullOrEmpty(source);
        }

        public static string GetFileNameWithoutExtension(this string source)
        {
            return Path.GetFileNameWithoutExtension(source);
        }

        public static string Truncate(this string value, int maxLength)
        {
            if (string.IsNullOrEmpty(value)) return value;
            return value.Length <= maxLength ? value : value.Substring(0, maxLength);
        }

        public static bool IsApproximate(this decimal first, decimal second, decimal margin)
        {
            return Math.Abs(first - second) <= margin;
        }
    }
}
