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
    public static class Watcher
    {
        private static readonly DbFacade dbFacade;

        static Watcher()
        {
            dbFacade = new DbFacade();
        }

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
                InitializeImport();
                Run();               
            }
            catch(Exception exc)
            {
                var save = true;

                if (exc.GetType() == typeof(MySqlException))
                    save = false;
                
                LogError($"Exception occured in {nameof(Watcher)}.OnCreated() with message {exc.Message}", save);
            }
            finally
            {
                LogInfo($"The number of database connections {DbFacade.Connections}");
                ApplicationState.Reset();
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


                    CleanupCPGReferenceMonthlyPlan(ref dsCPGReferenceMonthlyPlan);

                    CleanupTrackingResults(ref dsCPGPLResults, ref dsRetailerPLResults);

                    if (!ValidateHistoricalData(dsCpgpl, dsRetailerPL))
                        return;

                    if (!ValidateUniques(dsCpgpl, dsCpgProductHierarchy, dsCPGReferenceMonthlyPlan, dsMarketOverview, dsProductAttributes, dsRetailerPL, dsRetailerProductHierarchy, dsSellOutData, dsCPGPLResults, dsRetailerPLResults))
                        return;                                       

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

                    dbFacade.LoadFromStagingToCore (ApplicationState.ImportType.IsBase, ApplicationState.ImportType.IsMonthly, ApplicationState.ImportType.IsTracking);

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


        private static bool ValidateEANs (
            List<RetailerProductHierarchy> dsRetailerProductHierarchy,
            List<Cpgpl> dsCpgpl, 
            List<CPGReferenceMonthlyPlan> dsCPGReferenceMonthlyPlan, 
            DbFacade dbFacade)
        {
            ApplicationState.State = State.ValidatingEANs;

            if (ApplicationState.ImportType.IsBase)
            {
                var isValid = dsCpgpl.All(e => dsRetailerProductHierarchy.Exists(h => string.Equals(h.EAN, e.EAN)));

                if (!isValid)
                    LogError($"EANs cross-page validation has failed for {nameof(Cpgpl)} page");

                return isValid;
            }                

            if (ApplicationState.ImportType.IsMonthly)
            {
                var dbRetailerProductHierarchy = dbFacade.GetAll<RetailerProductHierarchy>();
                var isValid = dsCPGReferenceMonthlyPlan.All(e => dbRetailerProductHierarchy.Exists(h => string.Equals(h.EAN, e.EAN)));

                if (!isValid)
                    LogError($"EANs cross-page validation has failed for {nameof(CPGReferenceMonthlyPlan)} page");

                return isValid;

            }

            return true;
        }

        private static bool ValidateUniques(
            List<Cpgpl> dsCpgpl, 
            List<CpgProductHierarchy> dsCpgProductHierarchy, 
            List<CPGReferenceMonthlyPlan> dsCPGReferenceMonthlyPlan, 
            List<MarketOverview> dsMarketOverview,
            List<ProductAttributes> dsProductAttributes,
            List<RetailerPL> dsRetailerPL,
            List<RetailerProductHierarchy> dsRetailerProductHierarchy,
            List<SellOutData> dsSellOutData,
            List<CPGPLResults> dsCPGPLResults,
            List<RetailerPLResults> dsRetailerPLResults)
        {
            ApplicationState.State = State.ValidatingUniqueValues;

            LogInfo($"Validate For Unique Values");

            if (ApplicationState.ImportType.IsBase)
            {
                var dataSet1 = dsCpgpl.Select(x => new { x.Year, x.YearType, x.Retailer, x.Banner, x.Country, x.EAN }).Distinct();

                if (dataSet1.Count() != dsCpgpl.Count())
                {
                    LogError($"Year,YearType,Retailer,Banner,Country,EAN have duplicates in {nameof(Cpgpl)}");
                    return false;
                }

                var dataSet2 = dsCpgProductHierarchy.Select(x => new {x.EAN}).Distinct();

                if (dataSet2.Count() != dsCpgProductHierarchy.Count())
                {
                    LogError($"EAN has duplicates in {nameof(CpgProductHierarchy)}");
                    return false;
                }

                var dataSet4 = dsMarketOverview.Select(x => new { x.Year, x.YearType, x.CPG, x.Retailer, x.Banner, x.Country, x.CategoryGroup, x.NielsenCategory, x.Market, x.MarketDesc, x.Segment, x.SubSegment }).Distinct();

                if (dataSet4.Count() != dsMarketOverview.Count())
                {
                    LogError($"Year,YearType,CPG,Retailer,Banner,Country,CategoryGroup,NielsenCategory,Market,MarketDesc,Segment,SubSegment have duplicates in {nameof(MarketOverview)}");
                    return false;
                }
                
                var dataSet5 = dsProductAttributes.Select(x => new { x.EAN }).Distinct();

                if (dataSet5.Count() != dsProductAttributes.Count())
                {
                    LogError($"EAN has duplicates in {nameof(ProductAttributes)}");
                    return false;
                }

                var dataSet6 = dsRetailerPL.Select(x => new { x.Year, x.YearType, x.Retailer, x.Banner, x.Country, x.EAN }).Distinct();

                if (dataSet6.Count() != dsRetailerPL.Count())
                {
                    LogError($"Year,YearType,Retailer,Banner,Country,EAN have duplicates in {nameof(RetailerPL)}");
                    return false;
                }

                var dataSet7 = dsRetailerProductHierarchy.Select(x => new { x.Retailer, x.Banner, x.Country, x.EAN }).Distinct();

                if (dataSet7.Count() != dsRetailerProductHierarchy.Count())
                {
                    LogError($"Retailer,Banner,Country,EAN have duplicates in {nameof(RetailerProductHierarchy)}");
                    return false;
                }


                var dataSet8 = dsSellOutData.Select(x => new { x.Year, x.YearType, x.CPG, x.Retailer, x.Banner, x.Country, x.EAN }).Distinct();

                if (dataSet8.Count() != dsSellOutData.Count())
                {
                    LogError($" Year, YearType, CPG, Retailer, Banner, Country, EAN have duplicates in {nameof(SellOutData)}");
                    return false;
                }                
            }

            if (ApplicationState.ImportType.IsMonthly)
            {
                var dataSet3 = dsCPGReferenceMonthlyPlan.Select(x => new { x.Year, x.YearType, x.Retailer, x.Banner, x.Country, x.EAN }).Distinct();

                if (dataSet3.Count() != dsCPGReferenceMonthlyPlan.Count())
                {
                    LogError($"Year,YearType,Retailer,Banner,Country,EAN have duplicates in {nameof(CPGReferenceMonthlyPlan)}");
                    return false;
                }
            }

            if (ApplicationState.ImportType.IsTracking)
            {
                var dataSet9 = dsCPGPLResults.Select(x => new { x.Year, x.YearType, x.Month, x.Retailer, x.Banner, x.Country, x.EAN }).Distinct();

                if (dataSet9.Count() != dsCPGPLResults.Count())
                {
                    LogError($"Year, YearType, Month, Retailer, Banner, Country, EAN have duplicates in {nameof(CPGPLResults)}");
                    return false;
                }

                var dataSet10 = dsRetailerPLResults.Select(x => new { x.Year, x.YearType, x.Month, x.Retailer, x.Banner, x.Country, x.EAN }).Distinct();

                if (dataSet9.Count() != dsRetailerPLResults.Count())
                {
                    LogError($"Year, YearType, Month, Retailer, Banner, Country, EAN have duplicates in {nameof(RetailerPLResults)}");
                    return false;
                }
            }


            return true;
        }

        private static bool ValidateHistoricalData(List<Cpgpl> dsCpgpl, List<RetailerPL> dsRetailerPL )
        {
            if (ApplicationState.ImportType.IsBase)
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
            }

            return true;
        }

        private static void InitializeImport()
        {
            try
            {
                var fileWithoutExtension = ApplicationState.File.Name.GetFileNameWithoutExtension();

                var value = dbFacade.GetAll<ImportDetails>().FirstOrDefault(x => x.Uuid.ToString().Equals(fileWithoutExtension, StringComparison.OrdinalIgnoreCase));

                if (value == null)
                    throw new Exception("Failed to extract import information from fbn_import database.");

                switch (value.ImportType)
                {
                    case "full-import":
                        ApplicationState.ImportType.IsBase = true;
                        break;
                    case "monthly-plan":
                        ApplicationState.ImportType.IsMonthly = true;
                        break;
                    case "monthly-tracking":
                        ApplicationState.ImportType.IsTracking = true;
                        break;
                    default:
                        throw new Exception("There is no proper value for ImportType column. Allowed values are full-import, monthly-plan, monthly-tracking.");
                }

                ApplicationState.ImportDetails = value;
            }
            catch (Exception exc)
            {
                LogError($"Exception occured in {nameof(Watcher)}.GetImportInformation() with message {exc.Message}");
                throw exc;
            }            
        }

        private static void WaitForFile(FileSystemEventArgs arg)
        {            
            var attempts = 1;

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
                    Thread.Sleep(1000);

                    LogWarning("File copy in progress...");

                    if (++attempts >= 20)
                    {
                        string message = $"File cannot be copied. The number of {attempts} attempts is over. The file is too big or internet connection is slow.";
                        LogError($"Exception occured in {nameof(Watcher)}.WaitForFile() with message {message}");
                        throw new Exception(message);
                    }
                }
            }
        }


        private static void CleanupCPGReferenceMonthlyPlan(ref List<CPGReferenceMonthlyPlan> dsCPGReferenceMonthlyPlan)
        {
            if (ApplicationState.ImportType.IsMonthly)
            {                
                dsCPGReferenceMonthlyPlan = dsCPGReferenceMonthlyPlan.Where(x => x.Year == ApplicationState.ImportDetails.Year).ToList();             
            }
        }

        private static void CleanupTrackingResults(ref List<CPGPLResults> dsCPGPLResults, ref List<RetailerPLResults> dsRetailerPLResults)
        {
            if (ApplicationState.ImportType.IsTracking)
            {
                dsCPGPLResults = dsCPGPLResults.Where(x => x.Year == ApplicationState.ImportDetails.Year && x.Month == ApplicationState.ImportDetails.Month).ToList();
                dsRetailerPLResults = dsRetailerPLResults.Where(x => x.Year == ApplicationState.ImportDetails.Year && x.Month == ApplicationState.ImportDetails.Month).ToList();
            }
        }
    }
}
