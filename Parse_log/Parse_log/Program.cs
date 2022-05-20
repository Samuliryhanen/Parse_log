using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using static Parse_log.Json;

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
            Json json = new Json();
            Console.WriteLine(json.ProcessData("C://genretech/loginPuhdistus/Parse_log/test.txt"));
        }
    }
}