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
        public static List<T> Parse<T>(ExcelWorksheet worksheet) where T : IModel, new()
        {
            var sheet = typeof(T).Name;
            var data = new List<T>();

            int rowCount = worksheet.Dimension.Rows;
            int colCount = worksheet.Dimension.Columns;

            try
            {
                //Fetch data from spreadsheet file
                for (int row = 2; row <= rowCount; row++)
                {
                    T obj = new T();
                    var col = 1;

                    foreach (var prop in AttributeHelper.GetSortedProperties<T>())
                    {
                        object value = worksheet.Cells[row, col].Value;

                        try
                        {
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
                        catch (Exception innerException)
                        {
                            LogError($"Exception occured on type convert for property {prop.Name} with message: {innerException.Message}");
                            throw innerException;
                        }

                    }

                    if (!obj.IsEmpty())
                        data.Add(obj);
                }
            }
            catch (Exception outerException)
            {
                LogError($"Unhandled exception occured when parsing file with message: {outerException.Message}");
                throw outerException;
            }

            return data;

            //DbFacade db = new DbFacade();
            //db.Insert(data);

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
                var confSheets = AppSettings.GetInstance().mainsheets;
                var workookSheets = package.Workbook.Worksheets.Select(x => x.Name).ToArray();

                if (confSheets.All(x => workookSheets.Contains(x, StringComparer.OrdinalIgnoreCase)))
                    ApplicationState.HasRequiredSheets = true;

                if (workookSheets.Any(x => x.Contains("CPGReferenceMonthlyPlan", StringComparison.OrdinalIgnoreCase)))
                    ApplicationState.HasMonthlyPlanSheet = true;

                return ApplicationState.HasRequiredSheets || ApplicationState.HasMonthlyPlanSheet;
            }
        }

        public static bool IsPageValid<T>(ExcelWorksheet worksheet)
        {
            var sheet = typeof(T).Name;
            var response = true;

            try
            {
                int colCount = worksheet.Dimension.Columns;

                for (int col = 1; col <= colCount; col++)
                {
                    object cellValue = worksheet.Cells[1, col].Value;
                    if (cellValue == null)
                        continue;

                    string columnName = (string)cellValue;

                    if (columnName.EndsWith("(%)"))
                    {
                        columnName = columnName.TrimEnd("(%)".ToArray());
                    }

                    string propName = AttributeHelper.GetPropertyByKey<T>(col).Name;

                    if (!string.Equals(columnName.ReplaceSpace(), propName, StringComparison.OrdinalIgnoreCase))
                    {
                        LogInfo($"Column {propName} is expected but {columnName} found in sheet {sheet}.");
                        response = false;
                        break;
                    }
                    col++;
                }
            }
            catch (Exception exc)
            {
                LogError($"Unhandled exception occured in IsPageValid() method with message: {exc.Message}");
                return false;
            }           

            return response;
        }
    }
}
