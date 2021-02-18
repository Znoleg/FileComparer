using System;
using System.Collections.Generic;
using NLog;

namespace ComparerLevel3
{
    class Program
    {
        static (string originalPath, string modifiedPath) GetFilePaths(string possibleOriginal = null, string possibleModified = null)
        {
            if (possibleOriginal is null)
            {
                Console.WriteLine("Enter original file path: ");
                possibleOriginal = Console.ReadLine();
            }
            if (possibleModified is null)
            {
                Console.WriteLine("Enter modified file path: ");
                possibleModified = Console.ReadLine();
            }

            return (possibleOriginal, possibleModified);
        }

        static void Main(string[] args)
        {
            Logger logger = LogManager.GetCurrentClassLogger();

            string original = null, possible = null;
            try
            {
                if (args[0] != null) original = args[0];
                if (args[1] != null) possible = args[1];
            }
            catch { };
            var (originalPath, modifiedPath) = GetFilePaths(original, possible);

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
