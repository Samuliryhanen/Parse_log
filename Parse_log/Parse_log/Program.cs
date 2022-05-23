﻿/// <summary>
/// Program for reading UIpath-log messages and inserting them to more readable excel-file
/// @Author Samuli Ryhänen 25.05.2022
/// </summary>
namespace Parse_log
{
    internal static class Program
    {
        
        /// <summary>
        /// The main entry point for the application.
        /// </summary>
        [STAThread]
        static void Main(string[] args)
        {
            if(args.Length >= 1)
            {
                Log log = new Log();
                log.ProcessData(args[0]);

            }

        }
    }
}