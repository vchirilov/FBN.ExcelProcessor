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
        private readonly int BATCH = 100;
        private readonly MySqlConnection sqlConnection;
        private string ConnectionString { get; } = AppSettings.GetInstance().connectionString;

        public DbFacade()
        {
            sqlConnection = new MySqlConnection(ConnectionString);
        }       

        public void Insert<T>(List<T> items) where T : new()
        {
            LogInfo($"{typeof(T).Name} data loading into database...");
            
            var chunks = GetChunks(items, BATCH);

            //Build column list in INSERT statement
            var columns = string.Empty;
            
            foreach (var prop in AttributeHelper.GetSortedProperties<T>())
                columns += $",`{prop.Name}`";
            columns = columns.TrimStart(',');
                                    
            var dbTable = GetDbTable<T>();            

            ExecuteNonQuery($"TRUNCATE TABLE {dbTable};", $"Truncate has failed for table {dbTable}");

            using (sqlConnection)
            {                
                sqlConnection.Open();

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

                    using (MySqlCommand sqlCommand = new MySqlCommand(text.ToString(), sqlConnection))
                    {
                        try
                        {
                            sqlCommand.CommandType = CommandType.Text;
                            sqlCommand.ExecuteNonQuery();
                        }
                        catch (Exception exc)
                        {
                            LogInfo($"Insert has failed: {exc.Message}");
                        }
                    }
                }
            }

            LogInfo($"{typeof(T).Name} loaded.");
        }        

        public void ConvertToNull(string table, string column, string value)
        {
            ExecuteNonQuery($"UPDATE `{table}` SET {column} = NULL WHERE {column} = '{value}';");
        }

        public void ImportDataToCore(bool isMonthlyPlanOnly)
        {
            LogInfo("Importing data from staging database to core. Please wait...");

            ExecuteNonQuery("CALL fbn_core.import_data()", "Import data from staging to core has failed");

            LogInfo("Importing data from staging database to core finished succesfully.");
        }

        private void ExecuteNonQuery(string sqlStatement, string message = "SQL execution has failed")
        {
            using (sqlConnection)
            {
                sqlConnection.Open();

                using (MySqlCommand sqlCommand = new MySqlCommand(sqlStatement.ToString(), sqlConnection))
                {
                    try
                    {
                        sqlCommand.CommandType = CommandType.Text;
                        sqlCommand.ExecuteNonQuery();
                    }
                    catch (Exception exc)
                    {
                        LogInfo($"{message}: {exc.Message}");
                    }
                }
            }

        }        
    }
}
