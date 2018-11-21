using ExcelProcessor.Helpers;
using ExcelProcessor.Models;
using OfficeOpenXml;
using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Threading;

namespace ExcelProcessor
{
    public class Parser
    {
        public static void Run<T>() where T : IModel, new()
        {
            var sheet = typeof(T).Name;            
            var data = new List<T>();
            var attempts = 0;

            while (true)
            {                
                try
                {
                    using (ExcelPackage package = new ExcelPackage(FileManager.File))
                    {
                        Console.WriteLine();
                        Console.WriteLine($"{sheet} is being initialized...");

                        //If sheet doesn't exist, exit the method
                        if (package.Workbook.Worksheets.Where(x => x.Name.Equals(sheet, StringComparison.CurrentCultureIgnoreCase)).FirstOrDefault() == null)
                        {
                            Console.WriteLine($"Worksheet {sheet} is missing in input file.");
                            return;
                        }                        

                        using (ExcelWorksheet worksheet = package.Workbook.Worksheets[sheet])
                        {
                            int rowCount = worksheet.Dimension.Rows;
                            int colCount = worksheet.Dimension.Columns;                            

                            //Fetch data from spreadsheet file
                            for (int row = 2; row <= rowCount; row++)
                            {
                                dynamic obj = new T();
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

                    break;
                }
                catch (IOException)
                {
                    Thread.Sleep(500);
                    Console.WriteLine("File copy in process...");                    
                    if (++attempts >= 20) break;
                }                
            }

            DbFacade db = new DbFacade();
            db.Insert(data);

            if (new T() is CpgProductHierarchy)
            {
                var cpgHierarchy = CpgProductHierarchyTree.GetTreeNodes(data as List<CpgProductHierarchy>);
                db.Insert(cpgHierarchy);
            }
        }
    }
}
