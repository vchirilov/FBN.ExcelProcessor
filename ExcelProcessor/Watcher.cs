using ExcelProcessor.Helpers;
using ExcelProcessor.Models;
using System;
using System.Diagnostics;
using System.IO;
using static ExcelProcessor.Helpers.Utility;

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
            Stopwatch stopWatch = new Stopwatch();
            stopWatch.Start();

            Log($"File [{e.Name}] has been created.");
                        
            try
            {
                Parser.Run<ProductAttributes>();
                Parser.Run<MarketOverview>();
                Parser.Run<CpgProductHierarchy>();
                Parser.Run<SellOutData>();
                Parser.Run<RetailerPL>();
                Parser.Run<RetailerProductHierarchy>();
                Parser.Run<Cpgpl>();
                Parser.Run<CPGReferenceMonthlyPlan>();

                DbFacade dbCore = new DbFacade();
                dbCore.ImportDataToCore();
            }
            catch (Exception exc)
            {
                Log($"Exception has occured with message {exc.Message}");
            }

            stopWatch.Stop();
            TimeSpan ts = stopWatch.Elapsed;

            string elapsedTime = string.Format("{0:00}:{1:00}:{2:00}.{3:00}", ts.Hours, ts.Minutes, ts.Seconds, ts.Milliseconds / 10);
            Log($"Import duration: {elapsedTime}");


            FileManager.DeleteFile();
        }

        private static void OnDeleted(object sender, FileSystemEventArgs e)
        {            
            Log($"File [{e.Name}] has been deleted.");
        }


    }
}
