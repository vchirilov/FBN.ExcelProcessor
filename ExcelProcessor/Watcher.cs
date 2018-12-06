using ExcelProcessor.Helpers;
using ExcelProcessor.Models;
using OfficeOpenXml;
using System;
using System.Collections.Generic;
using System.Diagnostics;
using System.IO;
using System.Linq;
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
            ClearScreen();
            AddHeader();
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
            List<ProductAttributes> dsProductAttributes = null;
            List<MarketOverview> dsMarketOverview = null;
            List<CpgProductHierarchy> dsCpgProductHierarchy = null;
            List<SellOutData> dsSellOutData = null;
            List<RetailerPL> dsRetailerPL = null;
            List<RetailerProductHierarchy> dsRetailerProductHierarchy = null;
            List<Cpgpl> dsCpgpl = null;
            List<CPGReferenceMonthlyPlan> dsCPGReferenceMonthlyPlan = null;

            //Validate Workbook
            if (!Parser.IsWorkbookValid())
            {
                LogInfo("Workbook has failed validation.");            
                return;
            }

            using (var workbook = new Workbook())
            {
                Stopwatch stopWatch = new Stopwatch();
                stopWatch.Start();               

                try
                {
                    if (!ValidateAllPages(workbook))
                    {
                        LogInfo("Sheets validation has failed.");
                        return;
                    }

                    foreach (var worksheet in workbook.Worksheets)
                    {
                        if (worksheet.Key.Equals(nameof(ProductAttributes), StringComparison.OrdinalIgnoreCase))
                            dsProductAttributes = Parser.Parse<ProductAttributes>(worksheet.Value);

                        if (worksheet.Key.Equals(nameof(MarketOverview), StringComparison.OrdinalIgnoreCase))
                            dsMarketOverview = Parser.Parse<MarketOverview>(worksheet.Value);

                        if (worksheet.Key.Equals(nameof(CpgProductHierarchy), StringComparison.OrdinalIgnoreCase))
                            dsCpgProductHierarchy = Parser.Parse<CpgProductHierarchy>(worksheet.Value);

                        if (worksheet.Key.Equals(nameof(SellOutData), StringComparison.OrdinalIgnoreCase))
                            dsSellOutData = Parser.Parse<SellOutData>(worksheet.Value);

                        if (worksheet.Key.Equals(nameof(RetailerPL), StringComparison.OrdinalIgnoreCase))
                            dsRetailerPL = Parser.Parse<RetailerPL>(worksheet.Value);

                        if (worksheet.Key.Equals(nameof(RetailerProductHierarchy), StringComparison.OrdinalIgnoreCase))
                            dsRetailerProductHierarchy = Parser.Parse<RetailerProductHierarchy>(worksheet.Value);

                        if (worksheet.Key.Equals(nameof(Cpgpl), StringComparison.OrdinalIgnoreCase))
                            dsCpgpl = Parser.Parse<Cpgpl>(worksheet.Value);

                        if (worksheet.Key.Equals(nameof(CPGReferenceMonthlyPlan), StringComparison.OrdinalIgnoreCase))
                            dsCPGReferenceMonthlyPlan = Parser.Parse<CPGReferenceMonthlyPlan>(worksheet.Value);
                    }

                    if (!ValidateUniques(dsCpgpl, dsCpgProductHierarchy, dsCPGReferenceMonthlyPlan, dsMarketOverview, dsProductAttributes, dsRetailerPL, dsRetailerProductHierarchy, dsSellOutData))
                        return;

                    DbFacade dbFacade = new DbFacade();

                    if (!ValidateEANs(dsRetailerProductHierarchy, dsCpgpl, dsCPGReferenceMonthlyPlan, dbFacade))
                        return;                    

                    if (dsProductAttributes != null)
                        dbFacade.Insert(dsProductAttributes);

                    if (dsMarketOverview != null)
                        dbFacade.Insert(dsMarketOverview);

                    if (dsCpgProductHierarchy != null)
                        dbFacade.Insert(dsCpgProductHierarchy);

                    if (dsSellOutData != null)
                        dbFacade.Insert(dsSellOutData);

                    if (dsRetailerPL != null)
                        dbFacade.Insert(dsRetailerPL);

                    if (dsRetailerProductHierarchy != null)
                        dbFacade.Insert(dsRetailerProductHierarchy);

                    if (dsCpgpl != null)
                        dbFacade.Insert(dsCpgpl);

                    if (dsCPGReferenceMonthlyPlan != null)
                        dbFacade.Insert(dsCPGReferenceMonthlyPlan);


                    dbFacade.LoadFromStagingToCore
                        (ApplicationState.HasRequiredSheets,
                        ApplicationState.HasMonthlyPlanSheet);
                }
                catch (Exception exc)
                {
                    LogError($"Exception has occured with message {exc.Message}");
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
        }

        private static bool ValidateAllPages(Workbook workbook)
        {
            foreach (var worksheet in workbook.Worksheets)
            {
                if (worksheet.Key.Equals(nameof(ProductAttributes), StringComparison.OrdinalIgnoreCase) && !Parser.IsPageValid<ProductAttributes>(worksheet.Value))
                    return false;

                if (worksheet.Key.Equals(nameof(MarketOverview), StringComparison.OrdinalIgnoreCase) && !Parser.IsPageValid<MarketOverview>(worksheet.Value))
                    return false;

                if (worksheet.Key.Equals(nameof(CpgProductHierarchy), StringComparison.OrdinalIgnoreCase) && !Parser.IsPageValid<CpgProductHierarchy>(worksheet.Value))
                    return false;

                if (worksheet.Key.Equals(nameof(SellOutData), StringComparison.OrdinalIgnoreCase) && !Parser.IsPageValid<SellOutData>(worksheet.Value))
                    return false;

                if (worksheet.Key.Equals(nameof(RetailerPL), StringComparison.OrdinalIgnoreCase) && !Parser.IsPageValid<RetailerPL>(worksheet.Value))
                    return false;

                if (worksheet.Key.Equals(nameof(RetailerProductHierarchy), StringComparison.OrdinalIgnoreCase) && !Parser.IsPageValid<RetailerProductHierarchy>(worksheet.Value))
                    return false;

                if (worksheet.Key.Equals(nameof(Cpgpl), StringComparison.OrdinalIgnoreCase) && !Parser.IsPageValid<Cpgpl>(worksheet.Value))
                    return false;

                if (worksheet.Key.Equals(nameof(CPGReferenceMonthlyPlan), StringComparison.OrdinalIgnoreCase) && !Parser.IsPageValid<CPGReferenceMonthlyPlan>(worksheet.Value))
                    return false;
            }

            return true;
        }

        private static bool ValidateEANs (List<RetailerProductHierarchy> dsRetailerProductHierarchy, List<Cpgpl> dsCpgpl, List<CPGReferenceMonthlyPlan> dsCPGReferenceMonthlyPlan, DbFacade dbFacade)
        {
            if (ApplicationState.HasRequiredSheets && ApplicationState.HasMonthlyPlanSheet)
            {                
                var isValid1 = dsCpgpl.All(e => dsRetailerProductHierarchy.Exists(h => string.Equals(h.EAN, e.EAN)));
                var isValid2 = dsCPGReferenceMonthlyPlan.All(e => dsRetailerProductHierarchy.Exists(h => string.Equals(h.EAN, e.EAN)));

                if (!isValid1)
                    LogError($"EANs cross-page validation has failed for {nameof(Cpgpl)} page");

                if (!isValid2)
                    LogError($"EANs cross-page validation has failed for {nameof(CPGReferenceMonthlyPlan)} page");

                return isValid1 && isValid2;
            }

            if (ApplicationState.HasRequiredSheets)
            {
                var isValid = dsCpgpl.All(e => dsRetailerProductHierarchy.Exists(h => string.Equals(h.EAN, e.EAN)));

                if (!isValid)
                    LogError($"EANs cross-page validation has failed for {nameof(Cpgpl)} page");

                return isValid;
            }
                

            if (ApplicationState.HasMonthlyPlanSheet)
            {
                var dbRetailerProductHierarchy = dbFacade.GetAll<RetailerProductHierarchy>();
                var isValid = dsCPGReferenceMonthlyPlan.All(e => dbRetailerProductHierarchy.Exists(h => string.Equals(h.EAN, e.EAN)));

                if (!isValid)
                    LogError($"EANs cross-page validation has failed for {nameof(CPGReferenceMonthlyPlan)} page");

                return isValid;

            }

            return false;
        }

        private static bool ValidateUniques(
            List<Cpgpl> dsCpgpl, 
            List<CpgProductHierarchy> dsCpgProductHierarchy, 
            List<CPGReferenceMonthlyPlan> dsCPGReferenceMonthlyPlan, 
            List<MarketOverview> dsMarketOverview,
            List<ProductAttributes> dsProductAttributes,
            List<RetailerPL> dsRetailerPL,
            List<RetailerProductHierarchy> dsRetailerProductHierarchy,
            List<SellOutData> dsSellOutData)
        {
            if (ApplicationState.HasRequiredSheets)
            {
                var items1 = dsCpgpl.Select(x => new { x.Year, x.YearType, x.Retailer, x.Banner, x.Country, x.EAN }).Distinct();

                if (items1.Count() != dsCpgpl.Count())
                {
                    LogError($"Year,YearType,Retailer,Banner,Country,EAN have duplicates in {nameof(Cpgpl)}");
                    return false;
                }

                var items2 = dsCpgProductHierarchy.Select(x => new {x.EAN}).Distinct();

                if (items2.Count() != dsCpgProductHierarchy.Count())
                {
                    LogError($"EAN has duplicates in {nameof(CpgProductHierarchy)}");
                    return false;
                }

                var items4 = dsMarketOverview.Select(x => new { x.Year, x.YearType, x.CPG, x.Retailer, x.Banner, x.Country, x.CategoryGroup, x.NielsenCategory, x.Market, x.MarketDesc, x.Segment, x.SubSegment }).Distinct();

                if (items4.Count() != dsMarketOverview.Count())
                {
                    LogError($"Year,YearType,CPG,Retailer,Banner,Country,CategoryGroup,NielsenCategory,Market,MarketDesc,Segment,SubSegment have duplicates in {nameof(MarketOverview)}");
                    return false;
                }
                
                var items5 = dsProductAttributes.Select(x => new { x.EAN }).Distinct();

                if (items5.Count() != dsProductAttributes.Count())
                {
                    LogError($"EAN has duplicates in {nameof(ProductAttributes)}");
                    return false;
                }

                var items6 = dsRetailerPL.Select(x => new { x.Year, x.YearType, x.Retailer, x.Banner, x.Country, x.EAN }).Distinct();

                if (items6.Count() != dsRetailerPL.Count())
                {
                    LogError($"Year,YearType,Retailer,Banner,Country,EAN have duplicates in {nameof(RetailerPL)}");
                    return false;
                }

                var items7 = dsRetailerProductHierarchy.Select(x => new { x.Retailer, x.Banner, x.Country, x.EAN }).Distinct();

                if (items7.Count() != dsRetailerProductHierarchy.Count())
                {
                    LogError($"Retailer,Banner,Country,EAN have duplicates in {nameof(RetailerProductHierarchy)}");
                    return false;
                }


                var items8 = dsSellOutData.Select(x => new { x.Year, x.YearType, x.CPG, x.Retailer, x.Banner, x.Country, x.EAN }).Distinct();

                if (items8.Count() != dsSellOutData.Count())
                {
                    LogError($" Year, YearType, CPG, Retailer, Banner, Country, EAN have duplicates in {nameof(SellOutData)}");
                    return false;
                }                
            }

            if (ApplicationState.HasMonthlyPlanSheet)
            {
                var items3 = dsCPGReferenceMonthlyPlan.Select(x => new { x.Year, x.YearType, x.Retailer, x.Banner, x.Country, x.EAN }).Distinct();

                if (items3.Count() != dsCPGReferenceMonthlyPlan.Count())
                {
                    LogError($"Year,YearType,Retailer,Banner,Country,EAN have duplicates in {nameof(CPGReferenceMonthlyPlan)}");
                    return false;
                }
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
                    LogInfo("File copy in process...");
                    if (++attempts >= 20) break;
                }
            }
        }
    }
}
