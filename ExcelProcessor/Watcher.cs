using ExcelProcessor.Helpers;
using ExcelProcessor.Models;
using MySql.Data.MySqlClient;
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
            try
            {
                ClearScreen();
                AddHeader();
                WaitForFile(e);
                Run();               
            }
            catch(Exception exc)
            {
                if (exc.GetType() == typeof(MySqlException))
                    LogError($"Global Exception: {exc.Message}",false);
                else
                    LogError($"Global Exception: {exc.Message}");
            }
            finally
            {                
                ApplicationState.Reset();

                LogInfo($"The number of database connections {DbFacade.Connections}");

                FileManager.DeleteFile();                
            }            
        }

        private static void OnDeleted(object sender, FileSystemEventArgs e)
        {
            LogInfo($"File [{e.Name}] has been deleted.", false);
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
            List<CPGPLResults> dsCPGPLResults = null;
            List<RetailerPLResults> dsRetailerPLResults = null;

            //Validate authenticated user
            if (!ValidateAuthenticationUser())
                return;

            //Validate Workbook
            if (!Parser.IsWorkbookValid())
                return;

            using (var workbook = new Workbook())
            {
                try
                {
                    Stopwatch stopWatch = new Stopwatch();
                    stopWatch.Start();

                    //Validate Worksheets
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

                        if (worksheet.Key.Equals(nameof(CPGPLResults), StringComparison.OrdinalIgnoreCase))
                            dsCPGPLResults = Parser.Parse<CPGPLResults>(worksheet.Value);

                        if (worksheet.Key.Equals(nameof(RetailerPLResults), StringComparison.OrdinalIgnoreCase))
                            dsRetailerPLResults = Parser.Parse<RetailerPLResults>(worksheet.Value);
                    }

                    if (!ValidateHistoricalData(dsCpgpl, dsRetailerPL))
                        return;

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

                    if (dsCPGPLResults != null)
                        dbFacade.Insert(dsCPGPLResults);

                    if (dsRetailerPLResults != null)
                        dbFacade.Insert(dsRetailerPLResults);                  

                    dbFacade.LoadFromStagingToCore (ApplicationState.HasRequiredSheets, ApplicationState.HasMonthlyPlanSheet, ApplicationState.HasTrackingSheets);

                    stopWatch.Stop();
                    TimeSpan ts = stopWatch.Elapsed;

                    ApplicationState.State = State.Finished;

                    string elapsedTime = string.Format("{0:00}:{1:00}:{2:00}.{3:00}", ts.Hours, ts.Minutes, ts.Seconds, ts.Milliseconds / 10);
                    LogInfo($"Import duration: {elapsedTime}");
                }
                catch (Exception exc)
                {
                    LogError($"Exception occured in {nameof(Watcher)}.Run() with message {exc.Message}");
                    throw exc;
                }
            }                
        }
        
        private static bool ValidateAllPages(Workbook workbook)
        {
            ApplicationState.State = State.InitializingWorksheet;

            var model = typeof(IModel);

            var models = AppDomain.CurrentDomain.GetAssemblies()
                .SelectMany(s => s.GetTypes())
                .Where(p => model.IsAssignableFrom(p));

            var query =
                from type in models
                join worksheet in workbook.Worksheets on type.Name.ToLower() equals worksheet.Key.ToLower()
                select new { type, worksheet };

            foreach (var item in query)
            {
                if (!Parser.IsPageValid(item.type,item.worksheet.Value))
                    return false;                
            }

            return true;
        }

        private static bool ValidateAuthenticationUser()
        {
            string authUser = ApplicationState.File.Name.Substring2("__", "__");

            if (authUser.IsNullOrEmpty())
            {
                LogError($"Authenticated user cannot be identified.");
                return false;
            }
            else
            {
                ApplicationState.UserId = authUser;
                return true;
            }
        }

        private static bool ValidateEANs (
            List<RetailerProductHierarchy> dsRetailerProductHierarchy,
            List<Cpgpl> dsCpgpl, 
            List<CPGReferenceMonthlyPlan> dsCPGReferenceMonthlyPlan, 
            DbFacade dbFacade)
        {
            ApplicationState.State = State.ValidatingEANs;

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
            ApplicationState.State = State.ValidatingUniqueValues;

            LogInfo($"Validate For Unique Values");

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

        private static bool ValidateHistoricalData(List<Cpgpl> dsCpgpl, List<RetailerPL> dsRetailerPL )
        {
            ApplicationState.State = State.ValidatingHistoricalData;

            LogInfo($"Validate Historical Data");

            int currYear = DateTime.Now.Year;
            int year1 = dsCpgpl.Select(x => x.Year).Min();
            int year2 = dsCpgpl.Select(x => x.Year).Min();

            if (currYear == year1)
            {
                LogError($"{nameof(Cpgpl)} has no historical data");
                return false;
            }

            if (currYear == year2)
            {
                LogError($"{nameof(RetailerPL)} has no historical data");
                return false;
            }           

            return true;
        }

        private static void WaitForFile(FileSystemEventArgs arg)
        {            
            var attempts = 0;

            ApplicationState.State = State.CopyingFile;

            while (true)
            {
                try
                {
                    using (ExcelPackage package = new ExcelPackage(FileManager.File))
                    {
                        ApplicationState.File = FileManager.File;
                        LogInfo($"File [{arg.Name}] has been created.");
                    }

                    break;
                }
                catch (Exception)
                {
                    Thread.Sleep(500);
                    LogInfo("File copy in process...");
                    if (++attempts >= 20) break;
                }
            }
        }
    }
}
