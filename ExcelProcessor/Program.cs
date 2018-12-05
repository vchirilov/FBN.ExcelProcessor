using System;
using static ExcelProcessor.Helpers.Utility;

namespace ExcelProcessor
{
    class Program
    {
        static void Main(string[] args)
        {
            AddHeader();

            try
            {
                Watcher.WatchFile();
            }
            catch (Exception exc)
            {
                LogError($"Unhandled excpetion has occured with message: {exc.Message}");
            }
            
            Console.ReadKey();
        }

    }
}

