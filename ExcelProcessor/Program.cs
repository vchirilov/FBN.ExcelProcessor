using ExcelProcessor.Models;
using System;
using System.Collections.Generic;
using System.IO;

namespace ExcelProcessor
{
    class Program
    {
        static void Main(string[] args)
        {
            Console.WriteLine("Service has started...");
            Watcher.WatchFile();           

            //DbFacade db = new DbFacade();
            //List<Cpgpl> test = new List<Cpgpl>();

            //for(int i=0; i < 885; i++)
            //{
            //    test.Add(new Cpgpl { Year = i, YearType = "FY" + i, Retailer = "Test" + 1 });
            //}                       

            //db.Insert<Cpgpl>(test);

            Console.ReadKey();
        }
    }
}

