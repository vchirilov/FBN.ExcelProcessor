using ExcelProcessor.Models;
using System;
using System.IO;

namespace ExcelProcessor
{
    public class Watcher
    {
        public static void WatchFile()
        {
            FileSystemWatcher watcher = new FileSystemWatcher();     
            watcher.Path = FileManager.GetContainerFolder().Name;
            watcher.Filter = "*.xlsx";
            watcher.IncludeSubdirectories = false;
            watcher.EnableRaisingEvents = true;

            watcher.Created += OnCreated;
            //watcher.Created += (sender, e) => Console.WriteLine("File created");

            watcher.Deleted += OnDeleted;
            //watcher.Deleted += (sender, e) => Console.WriteLine("File deleted");
        }

        private static void OnCreated(object sender, FileSystemEventArgs e)
        {
            Console.WriteLine($"File [{e.Name}] has been created.");

            Parser.Run<Cpgpl>();
            Parser.Run<ProductAttributes>();
            Parser.Run<MarketOverview>();
            Parser.Run<CpgProductHierarchy>();

            FileManager.DeleteFile();
        }

        private static void OnDeleted(object sender, FileSystemEventArgs e)
        {
            Console.WriteLine($"File [{e.Name}] has been deleted.");
        }


    }
}
