/// <summary>
/// Program for reading UIpath-log messages and inserting them to more readable excel-file
/// @Author Samuli Ryhänen 24.05.2022
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
            Log log = new Log();
            string name = "";
            log.ProcessData(@"C:\genretech\loginPuhdistus\test\Meita1_tuhoa.txt");
            if (args.Length >= 1)
            {
                Console.WriteLine("Processing data...");
                try
                {
                    name = log.ProcessData(args[0]);
                    Console.WriteLine("Data in excel: " + name);
                }
                catch 
                {
                    Console.WriteLine("Error inserting data: " + name);
                }
            
            
            }
            
        }
    }
}