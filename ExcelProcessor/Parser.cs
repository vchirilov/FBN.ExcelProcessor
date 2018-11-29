using ExcelProcessor.Config;
using ExcelProcessor.Helpers;
using ExcelProcessor.Models;
using OfficeOpenXml;
using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Reflection;
using System.Threading;
using static ExcelProcessor.Helpers.Utility;

namespace ExcelProcessor
{
    public class Parser
    {
        public static void Run<T>() where T : IModel, new()
        {
            var sheet = typeof(T).Name;
            var data = new List<T>();

            using (ExcelPackage package = new ExcelPackage(FileManager.File))
            {                
                LogInfo($"{sheet} is being initialized...");

                using (ExcelWorksheet worksheet = package.Workbook.Worksheets[sheet])
                {
                    int rowCount = worksheet.Dimension.Rows;
                    int colCount = worksheet.Dimension.Columns;

                    //Fetch data from spreadsheet file
                    for (int row = 2; row <= rowCount; row++)
                    {
                        T obj = new T();
                        var col = 1;

                        foreach (var prop in AttributeHelper.GetSortedProperties<T>())
                        {
                            object value = worksheet.Cells[row, col].Value;

                            switch (prop.PropertyType.Name)
                            {
                                case "Int32":
                                    value = Convert.ToInt32(value);
                                    break;
                                case "Decimal":
                                    value = Convert.ToDecimal(value);
                                    break;
                                default:
                                    break;
                            }

                            typeof(T).GetProperty($"{prop.Name}").SetValue(obj, value);
                            col++;
                        }

                        if (!obj.IsEmpty())
                            data.Add(obj);
                    }
                }
            }

            DbFacade db = new DbFacade();
            db.Insert(data);

            //if (new T() is CpgProductHierarchy)
            //{
            //    var cpgHierarchy = CpgProductHierarchyTree.GetTreeNodes(data as List<CpgProductHierarchy>);

            //    db.Insert(cpgHierarchy);                
            //    db.ConvertToNull($"{GetDbTable<TreeNode>()}", "ParentId", "-1");
            //}
        }
        
        public static bool IsWorkbookValid()
        {
            LogInfo("Workook is being validated...");
            using (ExcelPackage package = new ExcelPackage(FileManager.File))
            {
                var confSheets = AppSettings.GetInstance().sheets;
                var workookSheets = package.Workbook.Worksheets.Select(x => x.Name).ToArray();

                if (confSheets.All(x => workookSheets.Contains(x, StringComparer.OrdinalIgnoreCase)))
                    ApplicationState.HasRequiredSheets = true;

                if (workookSheets.Any(x => x.Contains("CPGReferenceMonthlyPlan", StringComparison.OrdinalIgnoreCase)))
                    ApplicationState.HasMonthlyPlanSheet = true;

                return ApplicationState.HasRequiredSheets || ApplicationState.HasMonthlyPlanSheet;
            }
        }
    }
}
