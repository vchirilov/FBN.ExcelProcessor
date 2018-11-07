using System;
using System.Collections.Generic;
using System.IO;
using System.Text;
using System.Threading.Tasks;

namespace ExcelProcessor
{
    public class Watcher
    {
        public static void WatchFile()
        {
            FileSystemWatcher watcher = new FileSystemWatcher();     
            watcher.Path = FileManager.GetContainerFolder().Name;
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
            Console.WriteLine($"File [{e.Name}] has been created.");
            Parser.Run();
        }

        private static void OnDeleted(object sender, FileSystemEventArgs e)
        {
            Console.WriteLine($"File [{e.Name}] has been deleted.");
        }


    }
}
