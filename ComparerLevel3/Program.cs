using System;
using System.Collections.Generic;
using NLog;

namespace ComparerLevel3
{
    class Program
    {
        static (string originalPath, string modifiedPath) GetFilePaths()
        {
            Console.WriteLine("Enter original file path: ");
            var originalPath = Console.ReadLine();
            Console.WriteLine("Enter modified file path: ");
            var modifiedPath = Console.ReadLine();

            return (originalPath, modifiedPath);
        }

        static void Main(string[] args)
        {
            Logger logger = LogManager.GetCurrentClassLogger();

            var (originalPath, modifiedPath) = GetFilePaths();

            FileComparer fileComparer;
            while (true)
            {
                try
                {
                    fileComparer = new FileComparer(originalPath, modifiedPath, logger);
                }
                catch (Exception exception)
                {
                    logger.Error(exception, $"{exception.Message} with \"{originalPath}\", \"{modifiedPath}\" arguments");
                    
                    Console.WriteLine(exception.Message + "\nPlease provide new information!");
                    (originalPath, modifiedPath) = GetFilePaths();
                    continue;
                }

                break;
            }

            List<string> changes = (List<string>)fileComparer.GetDifference();
            foreach (string change in changes)
            {
                Console.WriteLine(change);
            }

            LogManager.Flush();

            Console.WriteLine("Press any key to exit.");
            Console.ReadKey();
        }
    }
}
