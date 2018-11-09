using ExcelProcessor.Config;
using ExcelProcessor.Models;
using MySql.Data.MySqlClient;
using System;
using System.Collections.Generic;
using System.Data;
using System.Reflection;
using System.Text;

namespace ExcelProcessor
{
    public class DbFacade
    {
        public void Insert<T>(List<T> items)
        {
            Console.WriteLine($"Sheet {typeof(T).Name} in progress...");

            var chunks = GetChunks(items, 100);

            //Build column list in INSERT statement
            var columns = string.Empty;
            foreach (var prop in typeof(T).GetProperties())
                columns += $",{prop.Name}";
            columns = columns.TrimStart(',');

            var modelAttr = (ModelAttribute)typeof(T).GetCustomAttribute(typeof(ModelAttribute));
            var connectionString = AppSettings.GetInstance().connectionString;

            using (var conn = new MySqlConnection(connectionString))
            {                
                conn.Open();

                foreach (var chunk in chunks)
                {
                    List<string> rows = new List<string>();

                    foreach (var item in chunk)
                    {
                        var pairs = DictionaryFromType(item);
                        var parameters = string.Empty;

                        foreach (var pair in pairs)
                        {
                            var val = pair.Value == null ? null : MySqlHelper.EscapeString(pair.Value.ToString());
                            parameters += $",'{val}'";
                        }
                            

                        rows.Add("(" + parameters.TrimStart(',') + ")");
                    }

                    var text = new StringBuilder($"INSERT INTO {modelAttr.Table} ({columns}) VALUES ");
                    text.Append(string.Join(",", rows));
                    text.Append(";");

                    using (MySqlCommand sqlCommand = new MySqlCommand(text.ToString(), conn))
                    {
                        try
                        {
                            sqlCommand.CommandType = CommandType.Text;
                            sqlCommand.ExecuteNonQuery();
                        }
                        catch (Exception exc)
                        {
                            Console.WriteLine($"Insert has failed: {exc.Message}");
                        }
                    }
                }
            }

            Console.WriteLine($"Sheet {typeof(T).Name} processed.");
        }

        private IEnumerable<List<T>> GetChunks<T>(List<T> source, int size = 100)
        {
            for (int i = 0; i < source.Count; i += size)
            {
                yield return source.GetRange(i, Math.Min(size, source.Count - i));
            }
        }

        private Dictionary<string, object> DictionaryFromType(object customType)
        {
            if (customType == null)
                return new Dictionary<string, object>();

            Type t = customType.GetType();
            PropertyInfo[] props = t.GetProperties();
            Dictionary<string, object> dict = new Dictionary<string, object>();

            foreach (PropertyInfo prp in props)
            {
                object value = prp.GetValue(customType, new object[] { });
                dict.Add(prp.Name, value);
            }
            return dict;
        }
    }
}
