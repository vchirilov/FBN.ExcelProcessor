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
            var chunks = GetChunks(items, 100);

            //Build column list in INSERT statement
            var columns = string.Empty;
            foreach (var prop in typeof(T).GetProperties())
                columns += $",{prop.Name}";
            columns = columns.TrimStart(',');

            var modelAttribute = (ModelAttribute)typeof(T).GetCustomAttribute(typeof(ModelAttribute));

            StringBuilder text = new StringBuilder($"INSERT INTO {modelAttribute.Table} ({columns}) VALUES ");

            using (var conn = new MySqlConnection("server=localhost;port=3306;database=fbn_staging;user=root;password=spartak_1"))
            {
                List<string> rows = new List<string>();

                foreach(var chunk in chunks)
                {
                    foreach (var item in chunk)
                    {
                        var pairs = DictionaryFromType(item);                        
                        var parameters = string.Empty;

                        foreach (var pair in pairs)
                            parameters += $",'{pair.Value}'";

                        rows.Add("(" + parameters.TrimStart(',') + ")");
                    }
                }
                
                text.Append(string.Join(",", rows));
                text.Append(";");

                conn.Open();
                using (MySqlCommand sqlCommand = new MySqlCommand(text.ToString(), conn))
                {
                    sqlCommand.CommandType = CommandType.Text;
                    sqlCommand.ExecuteNonQuery();
                }
            }
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
