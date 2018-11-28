using ExcelProcessor.Helpers;
using ExcelProcessor.Models;
using OfficeOpenXml;
using System;
using System.Diagnostics;
using System.IO;
using System.Threading;
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
            WaitForFile();

            LogInfo($"File [{e.Name}] has been created.");

            if (!Parser.IsWorkbookValid())
            {
                LogInfo("Workbook has failed validation.");
                return;
            }

            Stopwatch stopWatch = new Stopwatch();
            stopWatch.Start();           
                        
            try
            {
                var isMonthlyPlanOnly = ApplicationState.IsMonthlyPlanOnly;

                if (isMonthlyPlanOnly)
                {
                    Parser.Run<CPGReferenceMonthlyPlan>();
                }
                else
                {
                    Parser.Run<ProductAttributes>();
                    Parser.Run<MarketOverview>();
                    Parser.Run<CpgProductHierarchy>();
                    Parser.Run<SellOutData>();
                    Parser.Run<RetailerPL>();
                    Parser.Run<RetailerProductHierarchy>();
                    Parser.Run<Cpgpl>();
                    Parser.Run<CPGReferenceMonthlyPlan>();
                }                

                DbFacade dbCore = new DbFacade();
                dbCore.ImportDataToCore(isMonthlyPlanOnly);
            }
            catch (Exception exc)
            {
                LogInfo($"Exception has occured with message {exc.Message}");
            }

            stopWatch.Stop();
            TimeSpan ts = stopWatch.Elapsed;

            string elapsedTime = string.Format("{0:00}:{1:00}:{2:00}.{3:00}", ts.Hours, ts.Minutes, ts.Seconds, ts.Milliseconds / 10);
            LogInfo($"Import duration: {elapsedTime}");


            FileManager.DeleteFile();
        }

        private static void OnDeleted(object sender, FileSystemEventArgs e)
        {            
            LogInfo($"File [{e.Name}] has been deleted.");
        }

        private static void WaitForFile()
        {
            var attempts = 0;

            while (true)
            {
                try
                {
                    using (ExcelPackage package = new ExcelPackage(FileManager.File))
                    { }

                    break;
                }               

                catch (IOException)
                {
                    Thread.Sleep(500);
                    Console.WriteLine("File copy in process...");
                    if (++attempts >= 20) break;
                }
            }
        }
    }
}
