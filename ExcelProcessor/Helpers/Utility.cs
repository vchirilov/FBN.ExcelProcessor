using System;
using System.Collections.Generic;
using System.Reflection;
using System.Text;

namespace ExcelProcessor.Helpers
{
    public class Utility
    {
        public static string GetDbTable<T>()
        {
            var modelAttr = (ModelAttribute)typeof(T).GetCustomAttribute(typeof(ModelAttribute));

            return modelAttr.Table;
        }

        public static IEnumerable<List<T>> GetChunks<T>(List<T> source, int size = 10)
        {
            for (int i = 0; i < source.Count; i += size)
            {
                yield return source.GetRange(i, Math.Min(size, source.Count - i));
            }
        }

        public static Dictionary<string, object> DictionaryFromType(object instance)
        {
            if (instance == null)
                return new Dictionary<string, object>();

            var props = AttributeHelper.GetSortedProperties(instance);

            Dictionary<string, object> dict = new Dictionary<string, object>();

            foreach (PropertyInfo prop in props)
            {
                object value = prop.GetValue(instance, new object[] { });
                dict.Add(prop.Name, value);
            }
            return dict;
        }

        public static string Encode(string plainText)
        {
            var plainTextBytes = System.Text.Encoding.UTF8.GetBytes(plainText);
            return System.Convert.ToBase64String(plainTextBytes);
        }

        public static string Decode(string base64EncodedData)
        {
            var base64EncodedBytes = System.Convert.FromBase64String(base64EncodedData);
            return System.Text.Encoding.UTF8.GetString(base64EncodedBytes);
        }
    }
}
