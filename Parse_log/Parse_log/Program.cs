/// <summary>
/// Program for reading UIpath-log messages and inserting them to more readable excel-file
/// @Author Samuli Ryhänen 07.06.2022
/// @Genretech Oy
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

            if (args.Length >= 1)
            {

                string filePathIn = Path.GetFullPath(args[0]);
                string dirIn = Path.GetDirectoryName(filePathIn);
                string filePathOut;
                string dirOut;
                string excelFileName;

                bool exists = Directory.Exists(dirIn) && File.Exists(filePathIn);

                if (!exists)
                {
                    Console.WriteLine("Invalid path or file: " + "\nPath: " + dirIn + "\nFilename: " + Path.GetFileName(filePathIn));
                    return;
                }
                if ( args.Length == 2)
                {
                    filePathOut = Path.GetFullPath(args[1]);
                    dirOut = Path.GetDirectoryName(filePathOut);

                    exists = Directory.Exists(dirOut);

                    if (!exists)
                    {
                        Console.WriteLine("Invalid path or file: " + "\nPath: " + dirOut + "\nFilename: "+ Path.GetFileName(filePathOut));
                        return;
                    }
                    excelFileName = Path.GetFileNameWithoutExtension(filePathOut) + ".xlsx";
                }
                else
                {
                    dirOut = Path.GetDirectoryName(filePathIn);
                    excelFileName = Path.GetFileNameWithoutExtension(filePathIn) + ".xlsx";
                }

                Console.WriteLine("Processing data...");
                try
                {
                    string excelPath = Path.Combine(dirOut, excelFileName);
                    log.ProcessData(filePathIn, excelPath);
                    Console.WriteLine("Data in excel: " + excelPath);
                }
                catch(Exception e)
                {
                    Console.WriteLine("Error inserting data: " + e);
                }            
            }
            
        }
    }
}