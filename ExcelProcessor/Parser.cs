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
                    catch
                    {
                        throw ApplicationError.Create($"Page [{sheet}], column [{prop.Name}]. Can't convert value [{value}] to type {prop.PropertyType.Name}");
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

            LogInfo("Workbook is being validated...");

            try
            {
                using (ExcelPackage package = new ExcelPackage(ApplicationState.File))
                {
                    var mainConfiguredSheets = AppSettings.GetInstance().MainSheets;
                    var monthlyConfiguredSheet = AppSettings.GetInstance().MonthlySheet;
                    var trackingConfiguredSheets = AppSettings.GetInstance().TrackingSheets;

                    var worksheets = package.Workbook.Worksheets.Select(x => x.Name).ToArray();

                    if (ApplicationState.ImportType.IsBase && !mainConfiguredSheets.All(x => worksheets.Contains(x, StringComparer.OrdinalIgnoreCase)))
                    {
                        var missingSheets = string.Join(",", mainConfiguredSheets.Except(worksheets, StringComparer.OrdinalIgnoreCase).ToArray());
                        throw ApplicationError.Create($"Workbook is not valid for full-import type.Missing pages are {missingSheets}");
                    }                        

                    if (ApplicationState.ImportType.IsMonthly && !monthlyConfiguredSheet.All(x => worksheets.Contains(x, StringComparer.OrdinalIgnoreCase)))
                    {
                        var missingSheets = string.Join(",", monthlyConfiguredSheet.Except(worksheets, StringComparer.OrdinalIgnoreCase).ToArray());
                        throw ApplicationError.Create($"Workbook is not valid for monthly-plan type.Missing pages are {missingSheets}");
                    }                        

                    if (ApplicationState.ImportType.IsTracking && !trackingConfiguredSheets.All(x => worksheets.Contains(x, StringComparer.OrdinalIgnoreCase)))
                    {
                        var missingSheets = string.Join(",", trackingConfiguredSheets.Except(worksheets, StringComparer.OrdinalIgnoreCase).ToArray());
                        throw ApplicationError.Create($"Workbook is not valid for monthly-tracking type.Missing pages are {missingSheets}");
                    }                        
                }
            }
            catch (Exception exc)
            {
                throw exc;
            }            
        }
                
        public static void ValidatePage(Type type, ExcelWorksheet worksheet)
        {
            var sheet = type.Name;

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

                    string modelProperty = string.Empty;

                    try
                    {
                        modelProperty = AttributeHelper.GetPropertyByKey(type, col).Name;
                    }
                    catch {
                        continue;
                    }

                    if (!string.Equals(columnName.ReplaceSpace(), modelProperty, StringComparison.OrdinalIgnoreCase))
                        throw ApplicationError.Create($"Column {modelProperty} is expected but {columnName} found in sheet {sheet}.");
                }
            }
            catch (Exception exc)
            {
                throw exc;
            }
        }
    }
}
