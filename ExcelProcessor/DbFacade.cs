using ExcelProcessor.Config;
using ExcelProcessor.Helpers;
using ExcelProcessor.Models;
using MySql.Data.MySqlClient;
using System;
using System.Collections.Generic;
using System.Data;
using System.Linq;
using System.Reflection;
using System.Text;
using static ExcelProcessor.Helpers.Utility;

namespace ExcelProcessor
{
    public class DbFacade
    {
        const int BATCH = 100;

        public void Insert<T>(List<T> items) where T : new()
        {
            Console.WriteLine($"{typeof(T).Name} data loading into database...");
            
            var chunks = GetChunks(items, BATCH);

            //Build column list in INSERT statement
            var columns = string.Empty;
            
            foreach (var prop in AttributeHelper.GetSortedProperties<T>())
                columns += $",`{prop.Name}`";
            columns = columns.TrimStart(',');
                                    
            var dbTable = GetDbTable<T>();
            var connectionString = AppSettings.GetInstance().connectionString;

            ExecuteNonQuery($"TRUNCATE TABLE {dbTable};", $"Truncate has failed for table {dbTable}");

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

                    var text = new StringBuilder($"INSERT INTO {dbTable} ({columns}) VALUES ");
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

        public void ConvertToNull(string table, string column, string value)
        {
            ExecuteNonQuery($"UPDATE `{table}` SET {column} = NULL WHERE {column} = '{value}';");
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
    }
}
