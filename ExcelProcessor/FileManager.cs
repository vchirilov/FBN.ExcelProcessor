using System;
using System.Collections.Generic;
using System.IO;
using System.Text;

namespace ExcelProcessor
{
    public class FileManager
    {
        public static string GetContainerFolder() => Directory.CreateDirectory(Path.Combine(Environment.CurrentDirectory, "Container")).Name;

        public static string File { get;} = "Barilla.xlsx";

        public static bool IsFileLocked(FileInfo file)
        {
            FileStream stream = null;

            try
            {
                stream = file.Open(FileMode.Open,FileAccess.Read,FileShare.None);
            }
            catch (IOException)
            {
                return true;
            }
            finally
            {
                if (stream != null)
                    stream.Close();
            }

            //file is not locked
            return false;
        }
    }    
}
