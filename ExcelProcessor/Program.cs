using System;

namespace ExcelProcessor
{
    class Program
    {
        static void Main(string[] args)
        {
            WriteHeadLine();

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

        private static void WriteHeadLine()
        {
            Console.ForegroundColor = ConsoleColor.Green;
            Console.WriteLine("Service has started...");
            Console.ForegroundColor = ConsoleColor.Yellow;
            Console.WriteLine("WARNING: Press any key to close the service.");
            Console.ForegroundColor = ConsoleColor.White;
            Console.WriteLine("*********************************************");
            Console.WriteLine();
        }
    }
}

