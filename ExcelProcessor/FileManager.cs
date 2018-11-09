using System;
using System.IO;
using System.Linq;


namespace ExcelProcessor
{
    public class FileManager
    {
        public static DirectoryInfo GetContainerFolder() => Directory.CreateDirectory(Path.Combine(Environment.CurrentDirectory, "Container"));        
        public static FileInfo File
        {
            get
            {
                var folder = GetContainerFolder().Name;
                var filePath = Path.Combine(folder, FileName);
                return new FileInfo(filePath);
            }
            
        }
        public static void DeleteFile()
        {
            if (FileManager.IsFileLocked(File))
                Console.WriteLine("The file is locked");
            else
                File.Delete();
        }

        private static string FileName => GetContainerFolder().GetFiles().OrderByDescending(f => f.LastWriteTime).FirstOrDefault()?.Name;
        private static bool IsFileLocked(FileInfo file)
        {
            FileStream stream = null;

            try
            {
                stream = file.Open(FileMode.Open, FileAccess.Read, FileShare.None);
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
