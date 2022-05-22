using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using static Parse_log.Log;

namespace Parse_log
{
    internal static class Program
    {
        /// <summary>
        /// The main entry point for the application.
        /// </summary>
        [STAThread]
        static void Main()
        {
            Log log = new Log();
            Console.WriteLine(log.ProcessData("C://genretech/loginPuhdistus/test/short_test.txt"));
        }
    }
}