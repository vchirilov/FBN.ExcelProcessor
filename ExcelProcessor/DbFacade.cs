using ExcelProcessor.Config;
using ExcelProcessor.Helpers;
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
            Console.WriteLine($"{typeof(T).Name} data loading into database...");

            const int BATCH = 100;
            var chunks = GetChunks(items, BATCH);

            //Build column list in INSERT statement
            var columns = string.Empty;
            
            foreach (var prop in AttributeHelper.GetSortedProperties<T>())
                columns += $",{prop.Name}";
            columns = columns.TrimStart(',');


            var modelAttr = (ModelAttribute)typeof(T).GetCustomAttribute(typeof(ModelAttribute));
            var connectionString = AppSettings.GetInstance().connectionString;

            ExecuteNonQuery($"TRUNCATE TABLE {modelAttr.Table};", $"Truncate has failed for table {modelAttr.Table}");

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

            Console.WriteLine($"{typeof(T).Name} loaded.");
        }

        private void ExecuteNonQuery(string sqlStatement, string message = "SQL execution has failed")
        {
            var connectionString = AppSettings.GetInstance().connectionString;

            using (var conn = new MySqlConnection(connectionString))
            {
                conn.Open();

                using (MySqlCommand sqlCommand = new MySqlCommand(sqlStatement.ToString(), conn))
                {
                    try
                    {
                        sqlCommand.CommandType = CommandType.Text;
                        sqlCommand.ExecuteNonQuery();
                    }
                    catch (Exception exc)
                    {
                        Console.WriteLine($"{message}: {exc.Message}");
                    }
                }
            }

        }        

        private IEnumerable<List<T>> GetChunks<T>(List<T> source, int size = 10)
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

            var props = AttributeHelper.GetSortedProperties(customType);

            Dictionary<string, object> dict = new Dictionary<string, object>();

            foreach (PropertyInfo prop in props)
            {
                object value = prop.GetValue(customType, new object[] { });
                dict.Add(prop.Name, value);
            }
            return dict;
        }
    }
}
