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
        
        public static void LogInfo(string txt, bool save = true)
        {
            Console.WriteLine(txt);

            if (save)
                DbFacade.LogRecord(ApplicationState.State.ToString(), "Information", txt);            
        }

        public static void LogEmptyLine()
        {
            Console.WriteLine();
        }

        public static void LogApplicationError(ApplicationError err, bool save = true)
        {
            Console.ForegroundColor = ConsoleColor.Red;
            Console.WriteLine($"Error: {err.Text}. Import has failed.");
            Console.ForegroundColor = ConsoleColor.White;

            if (save)
                DbFacade.LogRecord(ApplicationState.State.ToString(), "Error", err.ToString(),err.GetStackTrace());
        }

        public static void LogError(string txt, bool save = true)
        {
            Console.ForegroundColor = ConsoleColor.Red;
            Console.WriteLine($"Error: {txt}");
            Console.ForegroundColor = ConsoleColor.White;

            if (save)
                DbFacade.LogRecord(ApplicationState.State.ToString(), "Error", txt);
        }

        public static void LogWarning(string txt, bool save = true)
        {
            Console.ForegroundColor = ConsoleColor.Yellow;
            Console.WriteLine($"WARNING: {txt}");
            Console.ForegroundColor = ConsoleColor.White;

            if (save)
                DbFacade.LogRecord(ApplicationState.State.ToString(), "Warning", txt);
        }

        public static void AddHeader1()
        {
            Console.ForegroundColor = ConsoleColor.Green;
            Console.WriteLine($"Service(version: 2019.01.31.01) has started at {DateTime.Now}...");
            Console.ForegroundColor = ConsoleColor.White;            
        }

        public static void AddHeader2()
        {            
            Console.ForegroundColor = ConsoleColor.Yellow;
            Console.WriteLine($"WARNING: Press [q+Enter] to close the service. Execution time: {DateTime.Now}");
            Console.ForegroundColor = ConsoleColor.White;
            Console.WriteLine("************************************************************************************");
        }

        public static void ClearScreen()
        {
            Console.Clear();
        }
    }
}
