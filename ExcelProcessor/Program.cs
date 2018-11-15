using System;

namespace ExcelProcessor
{
    class Program
    {
        static void Main(string[] args)
        {
            Console.WriteLine("Service has started...");
            try
            {
                Watcher.WatchFile();
            }
            catch (Exception exc)
            {
                Console.WriteLine($"Unhandled excpetion has occured with message: {exc.Message}");
            }
            
            Console.ReadKey();
        }
    }
}

