using ExcelProcessor.Config;
using OfficeOpenXml;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using static ExcelProcessor.Helpers.Utility;

namespace ExcelProcessor
{
    public class Workbook: IDisposable
    {
        public Workbook() => Initialize();        

        public Dictionary<string, ExcelWorksheet> Worksheets = new Dictionary<string, ExcelWorksheet>();

        private ExcelPackage package = null;
        
        public void Dispose()
        {
            foreach (var worksheet in Worksheets)
                worksheet.Value.Dispose();

            package.Dispose();
        }

        private void Initialize()
        {
            ApplicationState.State = State.InitializingWorksheet;

            package = new ExcelPackage(FileManager.File);

            if (ApplicationState.HasRequiredSheets)
            {
                foreach (var sheet in AppSettings.GetInstance().mainsheets)
                {
                    LogInfo($"Initializing page {sheet}...");
                    Worksheets.Add(sheet, package.Workbook.Worksheets[sheet]);
                }
            }

            if (ApplicationState.HasMonthlyPlanSheet)
            {
                foreach (var sheet in AppSettings.GetInstance().additionalsheets)
                {
                    LogInfo($"Initializing page {sheet}...");
                    Worksheets.Add(sheet, package.Workbook.Worksheets[sheet]);
                }
            }
        }
    }
}
