using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Text.Json;

/// <summary>
/// Class for reading ".txt"-file
/// </summary>
namespace Parse_log
{
    internal class Json
    {
        /// <summary>
        /// Process data from a file into an array
        /// </summary>
        /// <param name="pathname"></param>
        public string ProcessData(string pathname)
        {
            string status;
            try
            {
                string[] logs = ReadFile(pathname);
                status = AddToExcel(logs);
            }
            catch (Exception e)
            {
                status = "Error: " + (e.Message);
            }

            return status;
        }
        /// <summary>
        /// Write an array into an excel document
        /// </summary>
        /// <param name="logs"> array written</param>
        /// <returns> status code</returns>
        private string AddToExcel(string[] logs)
        {
            Excel excel = new Excel(@"test.xlsx", 1); // opens first worksheet of excel
            string excelCode;
            string[][] attributes = SeperateAttributes(logs);
            for(int i = 0; i< attributes.Length; i++)
            {
                for(int j = 0; j < attributes[i].Length; j++)
                {
                    string value;
                    if (attributes[i][j].Length > 1)
                    {

                        value = attributes[i][j];
                       // Console.WriteLine(value);
                        excelCode = SelectExcelCell(excel, value, i, j);
                        if (excelCode != "ok") return "Problem adding to excel: " + excelCode;
                    }
                }
            }
            excel.SaveAs(@"C:\genretech\loginPuhdistus\Parse_log\test.xlsx");
            excel.Close();
            return "ok";
        }
        private string SelectExcelCell(Excel excel, string value, int row, int column)
        {
            row++;
            column++;
            try
            {
                excel.Write(value, row, column);
            }
            catch(Exception e)
            {
                return e.ToString();
            }
            return "ok";
        }

        /// <summary>
        /// TODO: DOKUMENTOI
        /// </summary>
        /// <param name="logs"></param>
        /// <returns></returns>
        private string[][] SeperateAttributes(string[] logs)
        {
            string subs;
            string[][] attributes = new string[logs.Length][];
            string[] temp;
            for(int i = 0; i < logs.Length; i++)
            {
                subs = logs[i].Replace('"', ' '); // Remove the date-time string before actual JSON-notation
                temp = subs.Split(' ');
                string[] foo = Array.FindAll(temp, c => c.Length > 1); // copy everything thats longer than 1 char
                attributes[i] = foo;
            }
            return attributes; 
            
        }

        /// <summary>
        /// Read file from a given location and return every line of that file in an array
        /// </summary>
        /// <param name="pathname"> Given pathname</param>
        /// <returns>Array from the rows of the file</returns>
        private string[] ReadFile(string pathname)
        {
            string extension = Path.GetExtension(pathname);
            string[] file_text = { };

            if (extension == ".txt")
            {
                file_text = ReadTxt(pathname);
            }
            // Jos tiedostotyyppi on eri
            if (extension == ".json")
            {
                file_text = new string[] { "Json is not included yet" };
            }
            return file_text;
        }
        /// <summary>
        /// Reads a "txt"-file and returns an array of every row 
        /// </summary>
        /// <param name="pathname"> file location</param>
        /// <returns>Array, row as an index</returns>
        private string[] ReadTxt(string pathname)
        {
            // Read entire text file content in one string  
            string[] lines = File.ReadAllLines(pathname);
            return lines;
        }
    }
}