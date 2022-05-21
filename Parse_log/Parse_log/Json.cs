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

        ///TODO: aliohjelma pathin käsittelyyn

        /// <summary>
        /// Write an array into an excel document
        /// </summary>
        /// <param name="logs"> array written</param>
        /// <returns> status code</returns>
        static private string AddToExcel(string[] logs)
        {
            Excel excel = new Excel(@"short_test.xlsx", 1); // opens first worksheet of excel
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

                        excelCode = SelectExcelCell(excel, value, i, j);
                        if (excelCode != "ok") return "Problem adding to excel: " + excelCode;
                    }
                }
            }
            excel.SaveAs(@"C:\genretech\loginPuhdistus\test\short_test.xlsx");
            excel.Close();
            return "ok";
        }
        static private string SelectExcelCell(Excel excel, string value, int row, int column)
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
        /// iterate each line of logs[] and seperate the data from each line
        /// </summary>
        /// <param name="logs"></param>
        /// <returns>array with only attribute values</returns>
        static private string[][] SeperateAttributes(string[] logs)
        {
            //string subs;
            string[][] attributes = new string[logs.Length][];
            string[] arraySubs;
            for(int i = 0; i < logs.Length; i++)
            {
                //subs = logs[i].Replace('"', ' '); // Remove the date-time string before actual JSON-notation
                arraySubs = logs[i].Split('"', 2);
                // string[] temp = Array.FindAll(arraySubs, c => c.Length > 1); // copy everything thats longer than 1 char
                string[] temp = discardElements(arraySubs, true); //seperate function to remove 50% length from each individual array, and also to discard seperators such as ","
                attributes[i] = temp;                              // this case we need to have even elements
            }
            return attributes;
        }
        /// <summary>
        /// Function to discard even or odd index elements, and elements shorter than 1 length
        /// </summary>
        /// <param name="elements"> array to be manipulated</param>
        ///<param name = "even" > discard even or odd index elements</ param >
        /// <returns> new array</returns>
        static private string[] discardElements(string[] elements, bool even)
        {
            
            List<string> values = new List<string> {elements[0]};
            string[] json = elements[1].Split('"');
            int n;
            if (even)
            {
                n = 2;
            }
            else
            {
                n = 1;
            }
            int i = n;// start from index 2, because the first wanted attribute is on this index
            do
            {
                if (json[i].Length > 1 && i % n == 0)
                {
                    values.Add(json[i]);
                }
                i += 4;
            } while (i < json.Length);
            return values.ToArray();
        }
        /// <summary>
        /// Read file from a given location and return every line of that file in an array
        /// </summary>
        /// <param name="pathname"> Given pathname</param>
        /// <returns>Array from the rows of the file</returns>
        static private string[] ReadFile(string pathname)
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
        static private string[] ReadTxt(string pathname)
        {
            // Read entire text file content in one string  
            string[] lines = File.ReadAllLines(pathname);
            return lines;
        }
    }
}