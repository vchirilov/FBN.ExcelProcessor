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
                LogError($"Program exception: {exc.Message}", false);
            }
            
            Console.ReadKey();
        }

    }
}

