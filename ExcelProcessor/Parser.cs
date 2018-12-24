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
    public static class Parser
    {
        public static List<T> Parse<T>(ExcelWorksheet worksheet) where T : IModel, new()
        {
            var sheet = typeof(T).Name;
            var data = new List<T>();

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
                            case "String":
                                value = Convert.ToString(value);
                                break;
                            default:
                                break;
                        }

                        typeof(T).GetProperty($"{prop.Name}").SetValue(obj, value);
                        col++;
                    }
                    catch (Exception innerException)
                    {
                        LogError($"Exception occured in method {nameof(Parser)}.Parse<>() on type convert for property {sheet}.{prop.Name} with message: {innerException.Message}");
                        throw innerException;
                    }
                }

                if (!obj.IsEmpty())
                    data.Add(obj);
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
        
        public static void ValidateWorkbook()
        {
            ApplicationState.State = State.ValidatingWorkbook;

            LogInfo("Workook is being validated...");

            try
            {
                using (ExcelPackage package = new ExcelPackage(ApplicationState.File))
                {
                    var mainConfiguredSheets = AppSettings.GetInstance().mainsheets;
                    var monthlyConfiguredSheet = AppSettings.GetInstance().monthlysheet;
                    var trackingConfiguredSheets = AppSettings.GetInstance().trackingsheets;

                    var worksheets = package.Workbook.Worksheets.Select(x => x.Name).ToArray();

                    if (ApplicationState.ImportType.IsBase && !mainConfiguredSheets.All(x => worksheets.Contains(x, StringComparer.OrdinalIgnoreCase)))
                        throw new Exception("Workbook is not valid for full-import import type.");

                    if (ApplicationState.ImportType.IsMonthly && !monthlyConfiguredSheet.All(x => worksheets.Contains(x, StringComparer.OrdinalIgnoreCase)))
                        throw new Exception("Workbook is not valid for monthly-plan import type.");

                    if (ApplicationState.ImportType.IsTracking && !trackingConfiguredSheets.All(x => worksheets.Contains(x, StringComparer.OrdinalIgnoreCase)))
                        throw new Exception("Workbook is not valid for monthly-tracking import type.");
                }
            }
            catch (Exception exc)
            {
                LogError($"Exception occured in {nameof(Parser)}.IsWorkbookValid() with message {exc.Message}");
                throw exc;
            }            
        }
                
        public static bool IsPageValid(Type type, ExcelWorksheet worksheet)
        {
            var sheet = type.Name;
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

                    string modelProperty = AttributeHelper.GetPropertyByKey(type,col).Name;

                    if (!string.Equals(columnName.ReplaceSpace(), modelProperty, StringComparison.OrdinalIgnoreCase))
                    {
                        LogError($"Column {modelProperty} is expected but {columnName} found in sheet {sheet}.");
                        response = false;
                        break;
                    }
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
