using System;
using System.Collections.Generic;
using System.IO;
using System.Text;
using System.Threading.Tasks;

namespace ExcelProcessor
{
    public class Watcher
    {
        public static string Container { get; set; }

        public static void WatchFile()
        {
            var folder = FileManager.GetContainerFolder();

            FileSystemWatcher watcher = new FileSystemWatcher();
            Container = folder;
            watcher.Path = folder;
            watcher.Filter = "*.xlsx";
            watcher.IncludeSubdirectories = true;
            watcher.EnableRaisingEvents = true;

            watcher.Created += OnCreated;
            //watcher.Created += (sender, e) => Console.WriteLine("File created");

            watcher.Deleted += OnDeleted;
            //watcher.Deleted += (sender, e) => Console.WriteLine("File deleted");
        }

        private static void OnCreated(object sender, FileSystemEventArgs e)
        {
            Console.WriteLine("File created");
            Parser.Run();
        }

        private static void OnDeleted(object sender, FileSystemEventArgs e)
        {
            Console.WriteLine("File deleted");
        }


    }
}
