﻿using ExcelProcessor.Helpers;
using ExcelProcessor.Models;
using OfficeOpenXml;
using System;
using System.Collections.Generic;
using System.IO;
using System.Reflection;
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
                        using (ExcelWorksheet worksheet = package.Workbook.Worksheets[sheet])
                        {
                            int rowCount = worksheet.Dimension.Rows;
                            int colCount = worksheet.Dimension.Columns;                            

                            #region Validate Headers (Not possible as there are mismateches between database columns and Excel headers)
                            ////Validate headers
                            //var index = 1;
                            //foreach (var prop in typeof(T).GetProperties())
                            //{
                            //    if (!worksheet.Cells[1, index].Value.ToString().Compare(prop.Name))
                            //        throw new ArgumentException($"Header[1,{index}] is not {prop.Name}");
                            //    index++;
                            //} 
                            #endregion

                            //Fetch data from spreadsheet file
                            for (int row = 2; row <= rowCount; row++)
                            {
                                dynamic obj = new T();
                                var col = 1;

                                foreach (var prop in typeof(T).GetProperties())
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
        }
    }
}
