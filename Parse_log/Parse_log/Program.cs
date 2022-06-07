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
            args = new string[]{ "./subdir/short_test.txt", @"\uusi\uudempi\excel.xlsx"};
            
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
                        Directory.CreateDirectory(filePathOut);
                        Console.WriteLine("Path didn`t exists; New path created: " + filePathOut);
                    }
                    excelFileName = Path.GetFileNameWithoutExtension(filePathOut) + ".xlsx";
                }
                else
                {
                    filePathOut = filePathIn;
                    excelFileName = Path.GetFileNameWithoutExtension(filePathIn) + ".xlsx";
                }

                Console.WriteLine("Processing data...");
                try
                {
                    filePathOut = Path.Combine(filePathOut, excelFileName);
                    log.ProcessData(filePathIn, filePathOut);
                    Console.WriteLine("Data in excel: " + filePathOut);
                }
                catch(Exception e)
                {
                    Console.WriteLine("Error inserting data: " + e);
                }
            
            
            }
            
        }
    }
}