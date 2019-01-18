using System;
using static ExcelProcessor.Helpers.Utility;

namespace ExcelProcessor
{
    class Program
    {
        static void Main(string[] args)
        {
            try
            {
                AddHeader1();
                Watcher.WatchFile();
            }
            catch (Exception exc)
            {
                LogError($"Program exception: {exc.Message}", false);
            }

            while (true)
            {
                if (Console.ReadLine() == "q")
                    break;
            }
        }

    }
}

