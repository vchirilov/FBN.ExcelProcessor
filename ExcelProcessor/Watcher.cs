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
            WaitForFile(e);
            Run();
            FileManager.DeleteFile();
        }

        private static void OnDeleted(object sender, FileSystemEventArgs e)
        {
            LogInfo($"File [{e.Name}] has been deleted.");
        }

        private static void Run()
        {
            //Validate Workbook
            if (!Parser.IsWorkbookValid())
            {
                LogInfo("Workbook has failed validation.");            
                return;
            }

            if (!ValidateAllPages())
            {
                LogInfo("Sheets validation has failed.");
                return;
            }

            
            Stopwatch stopWatch = new Stopwatch();
            stopWatch.Start();

            try
            {
                if (ApplicationState.HasRequiredSheets)
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

                if (ApplicationState.HasMonthlyPlanSheet)
                    Parser.Run<CPGReferenceMonthlyPlan>();

                DbFacade dbCore = new DbFacade();
                dbCore.LoadFromStagingToCore
                    (ApplicationState.HasRequiredSheets,
                    ApplicationState.HasMonthlyPlanSheet);
            }
            catch (Exception exc)
            {
                LogInfo($"Exception has occured with message {exc.Message}");
            }
            finally
            {
                ApplicationState.Reset();
            }

            stopWatch.Stop();
            TimeSpan ts = stopWatch.Elapsed;

            string elapsedTime = string.Format("{0:00}:{1:00}:{2:00}.{3:00}", ts.Hours, ts.Minutes, ts.Seconds, ts.Milliseconds / 10);
            LogInfo($"Import duration: {elapsedTime}");
        }

        private static bool ValidateAllPages()
        {
            if (ApplicationState.HasRequiredSheets)
            {
                if (!Parser.IsPageValid<ProductAttributes>())
                    return false;
                if (!Parser.IsPageValid<MarketOverview>())
                    return false;
                if (!Parser.IsPageValid<CpgProductHierarchy>())
                    return false;
                if (!Parser.IsPageValid<SellOutData>())
                    return false;
                if (!Parser.IsPageValid<RetailerPL>())
                    return false;
                if (!Parser.IsPageValid<RetailerProductHierarchy>())
                    return false;
                if (!Parser.IsPageValid<Cpgpl>())
                    return false;
            }

            if (ApplicationState.HasMonthlyPlanSheet)
            {
                if (!Parser.IsPageValid<CPGReferenceMonthlyPlan>())
                    return false;
            }                

            return true;
        }


        private static void WaitForFile(FileSystemEventArgs arg)
        {
            var attempts = 0;

            while (true)
            {
                try
                {
                    using (ExcelPackage package = new ExcelPackage(FileManager.File))
                    {
                        LogInfo($"File [{arg.Name}] has been created.");
                    }

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
