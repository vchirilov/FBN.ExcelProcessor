using OfficeOpenXml;
using System;
using System.IO;
using System.Threading;

namespace ExcelProcessor
{
    public class Parser
    {
        public static void Run()
        {
            var ds = new System.Data.DataSet();
            var attempts = 0;

            var folder = FileManager.GetContainerFolder().Name;

            Console.WriteLine($"Selected file [{FileManager.File}]");

            var filePath = Path.Combine(folder, FileManager.File);   
            
            FileInfo file = new FileInfo(filePath);            

            while(true)
            {                
                try
                {
                    using (ExcelPackage package = new ExcelPackage(file))
                    {
                        using (ExcelWorksheet worksheet = package.Workbook.Worksheets["CPGPL"])
                        {
                            int rowCount = worksheet.Dimension.Rows;
                            int ColCount = worksheet.Dimension.Columns;

                            var rawText = string.Empty;

                            #region CodeForLaterUse
                            //for (int row = 1; row <= rowCount; row++)
                            //{
                            //    for (int col = 1; col <= ColCount; col++)
                            //    {
                            //        // This is just for demo purposes
                            //        rawText += worksheet.Cells[row, col].Value.ToString() + "\t";
                            //    }
                            //    rawText += "\r\n";
                            //} 
                            #endregion

                            for (int row = 1; row <= rowCount; row++)
                            {
                                if (worksheet.Cells[row, 12]?.Value != null)
                                    Console.WriteLine(worksheet.Cells[row, 12]?.Value?.ToString());
                            }
                        }
                    }

                    break;
                }
                catch (IOException exc)
                {
                    Thread.Sleep(500);
                    Console.WriteLine("File copy in process...");                    
                    if (++attempts >= 20) break;
                }                
            }            

            if (FileManager.IsFileLocked(file))
                Console.WriteLine("The file is locked");
            else
                file.Delete();
        }
    }
}
