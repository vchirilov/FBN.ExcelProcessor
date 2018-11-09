using ExcelProcessor.Config;
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
            Console.ReadKey();
        }
    }
}

