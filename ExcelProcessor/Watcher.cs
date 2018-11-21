using ExcelProcessor.Helpers;
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
                        
            try
            {
                Parser.Run<ProductAttributes>();
                Parser.Run<MarketOverview>();
                Parser.Run<CpgProductHierarchy>();
                Parser.Run<SellOutData>();
                Parser.Run<RetailerPL>();
                Parser.Run<RetailerProductHierarchy>();
                Parser.Run<Cpgpl>();
            }
            catch (Exception exc)
            {
                Console.WriteLine($"Exception has occured with message {exc.Message}");
            }
            

            FileManager.DeleteFile();
        }

        private static void OnDeleted(object sender, FileSystemEventArgs e)
        {
            Console.WriteLine();
            Console.WriteLine($"File [{e.Name}] has been deleted.");
        }


    }
}
