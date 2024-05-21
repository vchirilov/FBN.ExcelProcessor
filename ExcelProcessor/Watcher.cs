using ExcelProcessor.Config;
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
using System.Threading.Tasks;
using static ExcelProcessor.Helpers.Utility;

namespace ExcelProcessor
{
    public static class Watcher
    {
        private static readonly DbFacade dbFacade;
        private static readonly FileSystemWatcher watcher;

        static Watcher()
        {
            dbFacade = new DbFacade();
            watcher = new FileSystemWatcher();
            GC.KeepAlive(watcher);
        }

        public static void WatchFile()
        {
            try
            {                
                watcher.Path = FileManager.GetContainerFolder().Name;
                watcher.Filter = "*.xlsx";
                watcher.IncludeSubdirectories = false;
                watcher.EnableRaisingEvents = true;

                watcher.Created += OnCreated;
                watcher.Changed += OnChanged;
                watcher.Deleted += OnDeleted;
                watcher.Error += OnError;                
            }
            catch (Exception exc)
            {
                LogError($"Global exception: {exc.Message}", false);
            }            
        }

        private static void OnCreated(object sender, FileSystemEventArgs e)
        {
            try
            {
                //ClearScreen();
                AddHeader2();                
                WaitForFile(e);
                Run();                
            }
            catch(MySqlException exc)
            {
                LogError($"Database error: {exc.Message}", false);
            }
            catch (ApplicationError exc)
            {
                var appError = exc as ApplicationError;
                LogApplicationError(appError);                
            }
            catch (Exception exc)
            {
                LogError($"Unhandled exception:{exc.Message}");
            }
            finally
            {
                LogInfo($"The number of database connections {DbFacade.Instances}", false);

                ApplicationState.Reset();
                FileManager.DeleteFile();
            }            
        }

        private static void OnChanged(object sender, FileSystemEventArgs e)
        {
            LogWarning($"Watcher.OnChanged() event has occured", false);
        }

        private static void OnError(object sender, ErrorEventArgs e)
        {
            LogError($"Watcher.OnError() exception: {e.GetException().Message}", false);
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

            try
            {
                using (var workbook = new Workbook())
                {
                    Stopwatch stopWatch = new Stopwatch();
                    stopWatch.Start();

                    //Validate Workbook
                    Parser.ValidateWorkbook();

                    //Validate Worksheets
                    ValidateAllPages(workbook);

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

                    ValidateMonthlyPlan(dsCPGReferenceMonthlyPlan);

                    ValidateTrackingResults(dsCPGPLResults, dsRetailerPLResults);

                    ValidateHistoricalData(dsCpgpl, dsRetailerPL);

                    ValidateUniques(dsCpgpl, dsCpgProductHierarchy, dsCPGReferenceMonthlyPlan, dsMarketOverview, dsProductAttributes, dsRetailerPL, dsRetailerProductHierarchy, dsSellOutData, dsCPGPLResults, dsRetailerPLResults);

                    ValidateEANs(dsRetailerProductHierarchy, dsCpgpl, dsCPGReferenceMonthlyPlan, dbFacade);

                    ValidateSanityCheck(dsCpgpl, dsRetailerPL, dsCPGPLResults, dsRetailerPLResults);

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

                    dbFacade.LoadFromStagingToCore(ApplicationState.ImportType.IsBase, ApplicationState.ImportType.IsMonthly, ApplicationState.ImportType.IsTracking);

                    stopWatch.Stop();
                    TimeSpan ts = stopWatch.Elapsed;

                    ApplicationState.State = State.Finished;

                    string elapsedTime = string.Format("{0:00}:{1:00}:{2:00}.{3:00}", ts.Hours, ts.Minutes, ts.Seconds, ts.Milliseconds / 10);
                    LogInfo($"Import duration: {elapsedTime}");
                }
            }
            catch (Exception exc)
            {
                throw exc;
            }
                            
        }        

        private static void ValidateAllPages(Workbook workbook)
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
                Parser.ValidatePage(item.type, item.worksheet.Value);                
            }
        }


        private static void ValidateEANs (
            List<RetailerProductHierarchy> dsRetailerProductHierarchy,
            List<Cpgpl> dsCpgpl, 
            List<CPGReferenceMonthlyPlan> dsCPGReferenceMonthlyPlan, 
            DbFacade dbFacade)
        {
            if (ApplicationState.ImportType.IsBase)
            {
                ApplicationState.State = State.ValidatingEANs;
                LogInfo($"Validate EAN's");

                var isValid = dsCpgpl.All(e => dsRetailerProductHierarchy.Exists(h => string.Equals(h.EAN, e.EAN)));

                if (!isValid)
                    throw ApplicationError.Create($"EANs cross-page validation has failed for {nameof(Cpgpl)} page");
            }                

            if (ApplicationState.ImportType.IsMonthly)
            {
                ApplicationState.State = State.ValidatingEANs;
                LogInfo($"Validate EAN's");

                var dbRetailerProductHierarchy = dbFacade.GetAll<RetailerProductHierarchy>();
                var isValid = dsCPGReferenceMonthlyPlan.All(e => dbRetailerProductHierarchy.Exists(h => string.Equals(h.EAN, e.EAN)));

                if (!isValid)
                    throw ApplicationError.Create($"EANs cross-page validation has failed for {nameof(CPGReferenceMonthlyPlan)} page");
            }
        }

        private static void ValidateUniques(
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

            LogInfo($"Validate Unique Values");

            if (ApplicationState.ImportType.IsBase)
            {
                var dataSet1 = dsCpgpl.Select(x => new { x.Year, x.YearType, x.Retailer, x.Banner, x.Country, x.EAN }).Distinct();

                if (dataSet1.Count() != dsCpgpl.Count())
                    throw ApplicationError.Create($"Year,YearType,Retailer,Banner,Country,EAN have duplicates in {nameof(Cpgpl)}");

                var dataSet2 = dsCpgProductHierarchy.Select(x => new {x.EAN}).Distinct();

                if (dataSet2.Count() != dsCpgProductHierarchy.Count())
                    throw ApplicationError.Create($"EAN has duplicates in {nameof(CpgProductHierarchy)}");
                
                var dataSet4 = dsMarketOverview.Select(x => new { x.Year, x.YearType, x.CPG, x.Retailer, x.Banner, x.Country, x.CategoryGroup, x.NielsenCategory, x.Market, x.MarketDesc, x.Segment, x.SubSegment }).Distinct();

                if (dataSet4.Count() != dsMarketOverview.Count())
                    throw ApplicationError.Create($"Year,YearType,CPG,Retailer,Banner,Country,CategoryGroup,NielsenCategory,Market,MarketDesc,Segment,SubSegment have duplicates in {nameof(MarketOverview)}");
                                
                var dataSet5 = dsProductAttributes.Select(x => new { x.EAN }).Distinct();

                if (dataSet5.Count() != dsProductAttributes.Count())
                    throw ApplicationError.Create($"EAN has duplicates in {nameof(ProductAttributes)}");
                
                var dataSet6 = dsRetailerPL.Select(x => new { x.Year, x.YearType, x.Retailer, x.Banner, x.Country, x.EAN }).Distinct();

                if (dataSet6.Count() != dsRetailerPL.Count())
                    throw ApplicationError.Create($"Year,YearType,Retailer,Banner,Country,EAN have duplicates in {nameof(RetailerPL)}");

                var dataSet7 = dsRetailerProductHierarchy.Select(x => new { x.Retailer, x.Banner, x.Country, x.EAN }).Distinct();

                if (dataSet7.Count() != dsRetailerProductHierarchy.Count())
                    throw ApplicationError.Create($"Retailer,Banner,Country,EAN have duplicates in {nameof(RetailerProductHierarchy)}");

                var dataSet8 = dsSellOutData.Select(x => new { x.Year, x.YearType, x.CPG, x.Retailer, x.Banner, x.Country, x.EAN }).Distinct();

                if (dataSet8.Count() != dsSellOutData.Count())
                    throw ApplicationError.Create($" Year, YearType, CPG, Retailer, Banner, Country, EAN have duplicates in {nameof(SellOutData)}");
            }

            if (ApplicationState.ImportType.IsMonthly)
            {
                var dataSet3 = dsCPGReferenceMonthlyPlan.Select(x => new { x.Year, x.YearType, x.Retailer, x.Banner, x.Country, x.EAN }).Distinct();

                if (dataSet3.Count() != dsCPGReferenceMonthlyPlan.Count())
                    throw ApplicationError.Create($"Year,YearType,Retailer,Banner,Country,EAN have duplicates in {nameof(CPGReferenceMonthlyPlan)}");
            }

            if (ApplicationState.ImportType.IsTracking)
            {
                var dataSet9 = dsCPGPLResults.Select(x => new { x.Year, x.YearType, x.Month, x.Retailer, x.Banner, x.Country, x.EAN }).Distinct();

                if (dataSet9.Count() != dsCPGPLResults.Count())
                    throw ApplicationError.Create($"Year, YearType, Month, Retailer, Banner, Country, EAN have duplicates in {nameof(CPGPLResults)}");

                var dataSet10 = dsRetailerPLResults.Select(x => new { x.Year, x.YearType, x.Month, x.Retailer, x.Banner, x.Country, x.EAN }).Distinct();

                if (dataSet10.Count() != dsRetailerPLResults.Count())
                    throw ApplicationError.Create($"Year, YearType, Month, Retailer, Banner, Country, EAN have duplicates in {nameof(RetailerPLResults)}");
            }
        }

        private static void ValidateHistoricalData(List<Cpgpl> dsCpgpl, List<RetailerPL> dsRetailerPL )
        {
            if (ApplicationState.ImportType.IsBase)
            {
                ApplicationState.State = State.ValidatingHistoricalData;

                LogInfo($"Validate Historical Data");

                int currYear = DateTime.Now.Year;
                int minYear1 = dsCpgpl.Select(x => x.Year).Min();
                int minYear2 = dsRetailerPL.Select(x => x.Year).Min();

                if (currYear <= minYear1)
                    throw ApplicationError.Create($"{nameof(Cpgpl)} has no historical data");
                
                if (currYear <= minYear2)
                    throw ApplicationError.Create($"{nameof(RetailerPL)} has no historical data");
            }
        }
        private static void InitializeImport()
        {
            try
            {
                var fileWithoutExtension = ApplicationState.File.Name.GetFileNameWithoutExtension();

                var importDetails = dbFacade.GetAll<ImportDetails>().FirstOrDefault(x => x.Uuid.ToString().Equals(fileWithoutExtension, StringComparison.OrdinalIgnoreCase));
                var marginFromDb = dbFacade.GetAll<Configuration>().Where(x => x.ParameterName.ToLower() == "margin").Select(x => x.ParameterValue).FirstOrDefault();
                
                if (decimal.TryParse(marginFromDb, out decimal margin))
                    AppSettings.GetInstance().Margin = margin;

                if (importDetails == null)
                    throw ApplicationError.Create("Failed to extract import information from fbn_import database.");
                else
                    ApplicationState.ImportDetails = importDetails;


                switch (importDetails.ImportType)
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
                        throw ApplicationError.Create("There is no proper value for ImportType column. Allowed values are full-import, monthly-plan, monthly-tracking.");
                }                
            }
            catch (Exception exc)
            {
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
                        InitializeImport();
                        LogInfo($"File [{ApplicationState.ImportDetails.FileName}] has been uploaded");
                    }

                    break;
                }
                catch (Exception)
                {
                    Thread.Sleep(1000);

                    LogWarning("File copy in progress...");

                    if (++attempts >= 20)
                        throw ApplicationError.Create($"File cannot be copied. The number of {attempts} attempts is over. The file is too big or internet connection is slow.");
                }
            }
        }
        private static void ValidateMonthlyPlan(List<CPGReferenceMonthlyPlan> dsCPGReferenceMonthlyPlan)
        {
            if (ApplicationState.ImportType.IsMonthly)
            {
                var margin = AppSettings.GetInstance().Margin;

                if (!dsCPGReferenceMonthlyPlan.Exists(x => x.Year == ApplicationState.ImportDetails.Year))
                    throw ApplicationError.Create($"Selected year {ApplicationState.ImportDetails.Year} doesn't exist in input file.");

                if (dsCPGReferenceMonthlyPlan.Select(x => x.Year).Distinct().Count() > 1)
                    throw ApplicationError.Create($"There are too many years defined in the file.");
                
                if (!dsCPGReferenceMonthlyPlan.All(x => 1m.IsApproximate(x.Jan + x.Feb + x.Mar + x.Apr + x.May + x.Jun + x.Jul + x.Aug + x.Sep + x.Oct + x.Nov + x.Dec, margin)))
                    throw ApplicationError.Create($"SUM of monthly breakdown for one EAN must be 100%");
            }
        }
        private static void ValidateTrackingResults(List<CPGPLResults> dsCPGPLResults, List<RetailerPLResults> dsRetailerPLResults)
        {
            if (ApplicationState.ImportType.IsTracking)
            {
                var year = ApplicationState.ImportDetails.Year;
                var month = ApplicationState.ImportDetails.Month;

                if (!dsCPGPLResults.All(x => x.Year == year && x.Month == month))
                    throw ApplicationError.Create($"TrackingResults has failed validation for {nameof(CPGPLResults)}. [Year & Month] must have [{year} & {month}] values in input file.");

                if (!dsRetailerPLResults.All(x => x.Year == year && x.Month == month))
                    throw ApplicationError.Create($"TrackingResults has failed validation for {nameof(RetailerPLResults)}. [Year & Month] must have [{year} & {month}] values in input file.");

                //dsCPGPLResults = dsCPGPLResults.Where(x => x.Year == ApplicationState.ImportDetails.Year && x.Month == ApplicationState.ImportDetails.Month).ToList();
                //dsRetailerPLResults = dsRetailerPLResults.Where(x => x.Year == ApplicationState.ImportDetails.Year && x.Month == ApplicationState.ImportDetails.Month).ToList();
            }
        }
        private static void ValidateSanityCheck(List<Cpgpl> dsCpgpl, List<RetailerPL> dsRetailerPL, List<CPGPLResults> dsCPGPLResults, List<RetailerPLResults> dsRetailerPLResults)
        {
            var margin = AppSettings.GetInstance().Margin;

            ApplicationState.State = State.SanityCheck;

            LogInfo($"Validate sanity check with margin = {margin}");

            if (ApplicationState.ImportType.IsBase)
            {
                //Sanity check for sheet Cpgpl
                if (!dsCpgpl.All(x => x.SellInVolumeTotal.IsApproximate(x.SellInVolumePromo + x.SellInVolumeNonPromo, margin)))
                    throw ApplicationError.Create($"Formula [SellInVolumeTotal = SellInVolumePromo + SellInVolumeNonPromo] in page {nameof(Cpgpl)} is not satisfied");

                if (!dsCpgpl.All(x => x.TTSTotal.IsApproximate(x.ListPricePerUnit - x.NetNetPrice, margin)))
                    throw ApplicationError.Create($"Formula [TTSTotal = ListPricePerUnit – NetNetPrice] in page {nameof(Cpgpl)} is not satisfied");

                if (!dsCpgpl.All(x => x.ThreeNetPrice.IsApproximate((x.SellInVolumePromo * x.PromoPrice + x.SellInVolumeNonPromo * x.NetNetPrice) / x.SellInVolumeTotal, margin)))
                    throw ApplicationError.Create($"Formula [3NetPrice = (SellInVolumePromo*PromoPrice + SellInVolumeNonPromo*NetNetPrice) / SellInVolumeTotal] in page {nameof(Cpgpl)} is not satisfied");

                if (!dsCpgpl.All(x => x.CPPTotal.IsApproximate(x.NetNetPrice - x.ThreeNetPrice, margin)))
                    throw ApplicationError.Create($"Formula [CPPTotal = NetNetPrice-3NetPrice] in page {nameof(Cpgpl)} is not satisfied");

                if (!dsCpgpl.All(x => x.CPGProfitL1Total.IsApproximate(x.ThreeNetPrice - x.COGSTotal, margin)))
                    throw ApplicationError.Create($"Formula [CPGProfitL1total = 3NetPrice-COGSTotal] in page {nameof(Cpgpl)} is not satisfied");

                if (!dsCpgpl.All(x => x.CPGProfitL1NonPromo.IsApproximate(x.NetNetPrice - x.COGSTotal, margin)))
                    throw ApplicationError.Create($"Formula [CPGProfitL1NonPromo = NetNetPrice - COGSTotal] in page {nameof(Cpgpl)} is not satisfied");

                if (!dsCpgpl.All(x => x.CPGProfitL1Promo.IsApproximate(x.PromoPrice - x.COGSTotal, margin)))
                    throw ApplicationError.Create($"Formula [CPGProfitL1Promo = PromoPrice-COGSTotal] in page {nameof(Cpgpl)} is not satisfied");

                if (!dsCpgpl.All(x => x.CPGProfitL2Total.IsApproximate(x.CPGProfitL1Total - x.CPGCODBTotal, margin)))
                    throw ApplicationError.Create($"Formula [CPGProfitL2Total = CPGProfitL1Total-CODBTotal] in page {nameof(Cpgpl)} is not satisfied");

                if (!dsCpgpl.All(x => x.CPGProfitL3Total.IsApproximate(x.CPGProfitL2Total - x.CPGOverheadTotal, margin)))
                    throw ApplicationError.Create($"Formula [CPGProfitL3Total = CPGProfitL2Total-CPGOverheadTotal] in page {nameof(Cpgpl)} is not satisfied");

                if (!dsCpgpl.All(x => x.TTSOnTotal.IsApproximate(x.TTSOnConditional + x.TTSOnUnConditional, margin)))
                    throw ApplicationError.Create($"Formula [TTSOnTotal = TTSOnConditional + TTSOnUnConditional] in page {nameof(Cpgpl)} is not satisfied");

                if (!dsCpgpl.All(x => x.TTSOffTotal.IsApproximate(x.TTSOffConditional + x.TTSOffUnConditional, margin)))
                    throw ApplicationError.Create($"Formula [TTSOffTotal = TTSOffConditional + TTSOffUnConditional] in page {nameof(Cpgpl)} is not satisfied");

                if (!dsCpgpl.All(x => x.TTSOffTotal.IsApproximate(x.TTSOffConditional + x.TTSOffUnConditional, margin)))
                    throw ApplicationError.Create($"Formula [TTSOffTotal = TTSOffConditional + TTSOffUnConditional] in page {nameof(Cpgpl)} is not satisfied");

                if (!dsCpgpl.All(x => x.TTSTotal.IsApproximate(x.TTSOnTotal + x.TTSOffTotal, margin)))
                    throw ApplicationError.Create($"Formula [TTSTotal = TTSOnTotal + TTSOffTotal] in page {nameof(Cpgpl)} is not satisfied");

                if (!dsCpgpl.All(x => x.CPPTotal.IsApproximate(x.CPPOn + x.CPPOff, margin)))
                    throw ApplicationError.Create($"Formula [CPPTotal = CPPOn + CPPOff] in page {nameof(Cpgpl)} is not satisfied");


                //Sanity check for sheet RetailerPL
                if (!dsRetailerPL.All(x => x.SellOutPriceAverage.IsApproximate((x.SellOutVolumePromo * x.SellOutPricePromo + x.SellOutVolumeNonPromo * x.SellOutPriceNonPromo) / x.SellOutVolumeTotal, margin)))
                    throw ApplicationError.Create($"Formula [SellOutPriceAverage = (SellOutVolumePromo*SellOutPricePromo + SellOutVolumeNonPromo * SellOutPriceNonPromo) / SellOutVolumeTotal] in page {nameof(RetailerPL)} is not satisfied");

                if (!dsRetailerPL.All(x => x.RetailerProfitL1Total.IsApproximate(x.SellOutPriceAverage - x.COGSTotal, margin)))
                    throw ApplicationError.Create($"Formula [RetailerProfitL1Total = SellOutPriceAverrage – COGSTotal] in page {nameof(RetailerPL)} is not satisfied");

                if (!dsRetailerPL.All(x => x.RetailerProfitL1Promo.IsApproximate(x.SellOutPricePromo - x.COGSPromo, margin)))
                    throw ApplicationError.Create($"Formula [RetailerProfitL1Promo = SellOutPricePromo – COGSPromo] in page {nameof(RetailerPL)} is not satisfied");

                if (!dsRetailerPL.All(x => x.RetailerProfitL1NonPromo.IsApproximate(x.SellOutPriceNonPromo - x.COGSNonPromo, margin)))
                    throw ApplicationError.Create($"Formula [RetailerProfitL1NonPromo = SellOutPriceNonPromo – COGSNonPRomo] in page {nameof(RetailerPL)} is not satisfied");

                if (!dsRetailerPL.All(x => x.RetailerProfitL2Total.IsApproximate(x.RetailerProfitL1Total - x.RetailerCODBTotal, margin)))
                    throw ApplicationError.Create($"Formula [RetailerProfitL2Total = RetailerProfitL1Total - RetailerCODBTotal] in page {nameof(RetailerPL)} is not satisfied");

                if (!dsRetailerPL.All(x => x.RetailerProfitL3Total.IsApproximate(x.RetailerProfitL2Total - x.RetailerOverheadTotal, margin)))
                    throw ApplicationError.Create($"Formula [RetailerProfitL3Total = RetailerProfitL2Total - RetailerOverheadTotal] in page {nameof(RetailerPL)} is not satisfied");

                if (!dsRetailerPL.All(x => x.SellOutVolumeTotal.IsApproximate(x.SellOutVolumePromo + x.SellOutVolumeNonPromo, margin)))
                    throw ApplicationError.Create($"Formula [SellOutVolumeTotal = SellOutVolumePromo + SellOutVolumeNonPromo] in page {nameof(RetailerPL)} is not satisfied");
                                
                //Cross sanity check for sheets Gpgpl & RetailerPL
                RunPLCrossSheetValidation();
            }

            if (ApplicationState.ImportType.IsTracking)
            {
                //Sanity check for sheet CPGPLResults
                if (!dsCPGPLResults.All(x => x.SellInVolumeTotal.IsApproximate(x.SellInVolumePromo + x.SellInVolumeNonPromo, margin)))
                    throw ApplicationError.Create($"Formula [SellInVolumeTotal = SellInVolumePromo + SellInVolumeNonPromo] in page {nameof(CPGPLResults)} is not satisfied");

                if (!dsCPGPLResults.All(x => x.TTSTotal.IsApproximate(x.ListPricePerUnit - x.NetNetPrice, margin)))
                    throw ApplicationError.Create($"Formula [TTSTotal = ListPricePerUnit – NetNetPrice] in page {nameof(CPGPLResults)} is not satisfied");

                if (!dsCPGPLResults.All(x => x.ThreeNetPrice.IsApproximate((x.SellInVolumePromo * x.PromoPrice + x.SellInVolumeNonPromo * x.NetNetPrice) / x.SellInVolumeTotal, margin)))
                    throw ApplicationError.Create($"Formula [3NetPrice = (SellInVolumePromo*PromoPrice + SellInVolumeNonPromo*NetNetPrice) / SellInVolumeTotal] in page {nameof(CPGPLResults)} is not satisfied");

                if (!dsCPGPLResults.All(x => x.CPPTotal.IsApproximate(x.NetNetPrice - x.ThreeNetPrice, margin)))
                    throw ApplicationError.Create($"Formula [CPPTotal = NetNetPrice-3NetPrice] in page {nameof(CPGPLResults)} is not satisfied");

                if (!dsCPGPLResults.All(x => x.CPGProfitL1Total.IsApproximate(x.ThreeNetPrice - x.COGSTotal, margin)))
                    throw ApplicationError.Create($"Formula [CPGProfitL1total = 3NetPrice-COGSTotal] in page {nameof(CPGPLResults)} is not satisfied");

                if (!dsCPGPLResults.All(x => x.CPGProfitL1NonPromo.IsApproximate(x.NetNetPrice - x.COGSTotal, margin)))
                    throw ApplicationError.Create($"Formula [CPGProfitL1NonPromo = NetNetPrice - COGSTotal] in page {nameof(CPGPLResults)} is not satisfied");

                if (!dsCPGPLResults.All(x => x.CPGProfitL1Promo.IsApproximate(x.PromoPrice - x.COGSTotal, margin)))
                    throw ApplicationError.Create($"Formula [CPGProfitL1Promo = PromoPrice-COGSTotal] in page {nameof(CPGPLResults)} is not satisfied");

                if (!dsCPGPLResults.All(x => x.CPGProfitL2Total.IsApproximate(x.CPGProfitL1Total - x.CPGCODBTotal, margin)))
                    throw ApplicationError.Create($"Formula [CPGProfitL2Total = CPGProfitL1Total-CODBTotal] in page {nameof(CPGPLResults)} is not satisfied");

                if (!dsCPGPLResults.All(x => x.CPGProfitL3Total.IsApproximate(x.CPGProfitL2Total - x.CPGOverheadTotal, margin)))
                    throw ApplicationError.Create($"Formula [CPGProfitL3Total = CPGProfitL2Total-CPGOverheadTotal] in page {nameof(CPGPLResults)} is not satisfied");

                if (!dsCPGPLResults.All(x => x.TTSOnTotal.IsApproximate(x.TTSOnConditional + x.TTSOnUnConditional, margin)))
                    throw ApplicationError.Create($"Formula [TTSOnTotal = TTSOnConditional + TTSOnUnConditional] in page {nameof(CPGPLResults)} is not satisfied");

                if (!dsCPGPLResults.All(x => x.TTSOffTotal.IsApproximate(x.TTSOffConditional + x.TTSOffUnConditional, margin)))
                    throw ApplicationError.Create($"Formula [TTSOffTotal = TTSOffConditional + TTSOffUnConditional] in page {nameof(CPGPLResults)} is not satisfied");

                if (!dsCPGPLResults.All(x => x.TTSOffTotal.IsApproximate(x.TTSOffConditional + x.TTSOffUnConditional, margin)))
                    throw ApplicationError.Create($"Formula [TTSOffTotal = TTSOffConditional + TTSOffUnConditional] in page {nameof(CPGPLResults)} is not satisfied");

                if (!dsCPGPLResults.All(x => x.TTSTotal.IsApproximate(x.TTSOnTotal + x.TTSOffTotal, margin)))
                    throw ApplicationError.Create($"Formula [TTSTotal = TTSOnTotal + TTSOffTotal] in page {nameof(CPGPLResults)} is not satisfied");

                if (!dsCPGPLResults.All(x => x.CPPTotal.IsApproximate(x.CPPOn + x.CPPOff, margin)))
                    throw ApplicationError.Create($"Formula [CPPTotal = CPPOn + CPPOff] in page {nameof(CPGPLResults)} is not satisfied");


                //Sanity check for sheet RetailerPLResults
                if (!dsRetailerPLResults.All(x => x.SellOutPriceAverage.IsApproximate((x.SellOutVolumePromo*x.SellOutPricePromo + x.SellOutVolumeNonPromo*x.SellOutPriceNonPromo)/x.SellOutVolumeTotal, margin)))
                    throw ApplicationError.Create($"Formula [SellOutPriceAverage = (SellOutVolumePromo*SellOutPricePromo + SellOutVolumeNonPromo * SellOutPriceNonPromo) / SellOutVolumeTotal] in page {nameof(RetailerPLResults)} is not satisfied");

                if (!dsRetailerPLResults.All(x => x.RetailerProfitL1Total.IsApproximate(x.SellOutPriceAverage - x.COGSTotal, margin)))
                    throw ApplicationError.Create($"Formula [RetailerProfitL1Total = SellOutPriceAverrage – COGSTotal] in page {nameof(RetailerPLResults)} is not satisfied");

                if (!dsRetailerPLResults.All(x => x.RetailerProfitL1Promo.IsApproximate(x.SellOutPricePromo - x.COGSPromo, margin)))
                    throw ApplicationError.Create($"Formula [RetailerProfitL1Promo = SellOutPricePromo – COGSPromo] in page {nameof(RetailerPLResults)} is not satisfied");

                if (!dsRetailerPLResults.All(x => x.RetailerProfitL1NonPromo.IsApproximate(x.SellOutPriceNonPromo - x.COGSNonPromo, margin)))
                    throw ApplicationError.Create($"Formula [RetailerProfitL1NonPromo = SellOutPriceNonPromo – COGSNonPRomo] in page {nameof(RetailerPLResults)} is not satisfied");

                if (!dsRetailerPLResults.All(x => x.RetailerProfitL2Total.IsApproximate(x.RetailerProfitL1Total - x.RetailerCODBTotal, margin)))
                    throw ApplicationError.Create($"Formula [RetailerProfitL2Total = RetailerProfitL1Total - RetailerCODBTotal] in page {nameof(RetailerPLResults)} is not satisfied");

                if (!dsRetailerPLResults.All(x => x.RetailerProfitL3Total.IsApproximate(x.RetailerProfitL2Total - x.RetailerOverheadTotal, margin)))
                    throw ApplicationError.Create($"Formula [RetailerProfitL3Total = RetailerProfitL2Total - RetailerOverheadTotal] in page {nameof(RetailerPLResults)} is not satisfied");

                if (!dsRetailerPLResults.All(x => x.SellOutVolumeTotal.IsApproximate(x.SellOutVolumePromo + x.SellOutVolumeNonPromo, margin)))
                    throw ApplicationError.Create($"Formula [SellOutVolumeTotal = SellOutVolumePromo + SellOutVolumeNonPromo] in page {nameof(RetailerPLResults)} is not satisfied");
                                
                //Cross sanity check for sheets CPGPLResults & RetailerPLResults
                RunTrackingCrossSheetValidation();
            }

            void RunPLCrossSheetValidation()
            {
                var query =
                    from x in dsCpgpl
                    join y in dsRetailerPL on x.Year equals y.Year
                    where x.EAN.ToLower() == y.EAN.ToLower()
                    select new { x.ThreeNetPrice, x.PromoPrice, x.NetNetPrice, y.COGSTotal, y.COGSPromo, y.COGSNonPromo };

                var queryResult = query.ToList();

                if (!queryResult.All(x => x.ThreeNetPrice.IsApproximate(x.COGSTotal, margin)))
                    throw ApplicationError.Create($"Formula [CPG.ThreeNetPrice =  Retailer.COGSTotal] between pages {nameof(Cpgpl)} & {nameof(RetailerPL)} is not satisfied");

                if (!queryResult.All(x => x.PromoPrice.IsApproximate(x.COGSPromo, margin)))
                    throw ApplicationError.Create($"Formula [CPG.PromoPrice = Retailer. COGSPromo] between pages {nameof(Cpgpl)} & {nameof(RetailerPL)} is not satisfied");

                if (!queryResult.All(x => x.NetNetPrice.IsApproximate(x.COGSNonPromo, margin)))
                    throw ApplicationError.Create($"Formula [CPG.NetNetPrice = Retailer.COGSNonPromo] between pages {nameof(Cpgpl)} & {nameof(RetailerPL)} is not satisfied");
            }

            void RunTrackingCrossSheetValidation()
            {
                var query =
                    from x in dsCPGPLResults
                    join y in dsRetailerPLResults on x.Year equals y.Year
                    where x.EAN.ToLower() == y.EAN.ToLower()
                    select new { x.ThreeNetPrice, x.PromoPrice, x.NetNetPrice, y.COGSTotal, y.COGSPromo, y.COGSNonPromo };

                var queryResult = query.ToList();

                if (!queryResult.All(x => x.ThreeNetPrice.IsApproximate(x.COGSTotal, margin)))
                    throw ApplicationError.Create($"Formula [CPG.ThreeNetPrice =  Retailer.COGSTotal] between pages {nameof(CPGPLResults)} & {nameof(RetailerPLResults)} is not satisfied");

                if (!queryResult.All(x => x.PromoPrice.IsApproximate(x.COGSPromo, margin)))
                    throw ApplicationError.Create($"Formula [CPG.PromoPrice = Retailer. COGSPromo] between pages {nameof(CPGPLResults)} & {nameof(RetailerPLResults)} is not satisfied");

                if (!queryResult.All(x => x.NetNetPrice.IsApproximate(x.COGSNonPromo, margin)))
                    throw ApplicationError.Create($"Formula [CPG.NetNetPrice = Retailer.COGSNonPromo] between pages {nameof(CPGPLResults)} & {nameof(RetailerPLResults)} is not satisfied");
            }
        }
    }
}
