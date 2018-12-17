using Dapper;
using ExcelProcessor.Config;
using ExcelProcessor.Helpers;
using MySql.Data.MySqlClient;
using System;
using System.Collections.Generic;
using System.Data;
using System.Linq;
using System.Text;
using System.Threading;
using static ExcelProcessor.Helpers.Utility;

namespace ExcelProcessor
{
    public class DbFacade
    {
        public static int Connections = 0;
        private readonly int BATCH = 100;
        private readonly MySqlConnection sqlConnection;        
        private string ConnectionString { get; } = AppSettings.GetInstance().connectionString;

        public DbFacade()
        {
            sqlConnection = new MySqlConnection(ConnectionString);
            Interlocked.Increment(ref Connections);
        }       

        public void Insert<T>(List<T> items) where T : new()
        {
            ApplicationState.State = State.Loading;            
            
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
                            LogError($"Insert has failed: {exc.Message}");
                        }
                    }
                }
            }

            LogInfo($"{typeof(T).Name} loaded.");
        }

        public List<T> GetAll<T>()
        {
            using (sqlConnection)
            {
                sqlConnection.Open();

                var sql = $"SELECT * FROM fbn_staging.{GetDbTable<T>()}";

                return sqlConnection.Query<T>(sql).ToList();
            }
        }

        public void ConvertToNull(string table, string column, string value)
        {
            ExecuteNonQuery($"UPDATE `{table}` SET {column} = NULL WHERE {column} = '{value}';");
        }

        public static void LogRecord(string stage, string status, string message)
        {
            DbFacade db = new DbFacade();
            db.ExecuteNonQuery($"INSERT INTO fbn_logs.logs (`UserId`,`FileName`,`Stage`,`Status`,`Message`) VALUES ('{ApplicationState.UserId}','{ApplicationState.FileName}','{stage}','{status}','{message}');");
        }

        public void LoadFromStagingToCore(bool includeRequired, bool includeMonthlyPlan)
        {
            LogInfo("Importing data from staging database to core. Please wait...");

            ExecuteNonQuery($"CALL fbn_core.import_data_extended('{includeRequired.ToString()}','{includeMonthlyPlan.ToString()}')", "Import data from staging to core has failed");

            LogInfo("Importing data from staging database to core finished succesfully.");
        }

        private void ExecuteNonQuery(string sqlStatement, string message = "SQL execution has failed")
        {
            using (sqlConnection)
            {
                sqlConnection.Open();
                var sqlTransaction = sqlConnection.BeginTransaction();

                using (MySqlCommand sqlCommand = new MySqlCommand(sqlStatement.ToString(), sqlConnection, sqlTransaction))
                {
                    try
                    {
                        sqlCommand.CommandType = CommandType.Text;
                        sqlCommand.ExecuteNonQuery();
                        sqlTransaction.Commit();
                    }
                    catch (Exception exc)
                    {
                        sqlTransaction.Rollback();
                        LogError($"Exception has occured in method {nameof(DbFacade)}.ExecuteNonQuery() with message {exc.Message}");
                        throw exc;
                    }
                }

                Close(sqlConnection);
            }
        }

        private void Close(MySqlConnection sqlConnection)
        {
            if (sqlConnection != null && sqlConnection.State != ConnectionState.Closed)
                sqlConnection.Close();
        }

        ~DbFacade()
        {
            Interlocked.Decrement(ref Connections);
        }
    }    
}
