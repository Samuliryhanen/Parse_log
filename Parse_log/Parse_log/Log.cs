using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

/// <summary>
/// Class for reading ".txt"-file
/// </summary>
namespace Parse_log
{
    internal class Log
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



        /// TODO: aliohjelma pathin käsittelyyn



        /// <summary>
        /// Write an array into an excel document
        /// </summary>
        /// <param name="logs"> array written</param>
        /// <returns> status code</returns>
        static private string AddToExcel(string[] logs)
        {
            Excel excel = new Excel(@"short_test.xlsx", 1); // opens first worksheet of excel
            
            List<Dictionary<string, string>> logsList = SeperateAttributes(logs);
            for(int i = 0; i< logsList.Count; i++)
            {
                try
                {
                    AddRow(excel, logsList[i], i + 1);
                }
                catch(Exception e)
                {
                    return "Something went wrong: " + e; 
                }
            }
            excel.SaveAs(@"C:\genretech\loginPuhdistus\test\short_test.xlsx");
            excel.Close();
            return "ok";
        }
        static private void AddRow(Excel excel, Dictionary<string, string> logLine, int row)
        {
            
            var keys = logLine.Keys;
            foreach(string key in keys){
                switch (key)
                {
                    case "timeStamp":
                        excel.Write(logLine["timeStamp"], row, 1);
                        // TODO: tähän funktio, jolla valitaan väri
                        int colour = System.Drawing.ColorTranslator.ToOle(System.Drawing.Color.Red);
                        excel.CellColor(row, 1, colour);
                        break;
                    case "level":
                        excel.Write(logLine["level"], row, 2);
                        break;
                    case "processName":
                        excel.Write(logLine["processName"], row, 3);
                        break;
                    case "message":
                        excel.Write(logLine["message"], row, 4);
                        break;
                    case "fileName":
                        excel.Write(logLine["fileName"], row, 5);
                        break;
                    case "processVersion":
                        excel.Write(logLine["processVersion"], row, 6);
                        break;
                    case "robotName":
                        excel.Write(logLine["robotName"], row, 7);
                        break;
                        
                    default:
                        // TODO:Loppujen arvojen syöttö soluihin?
                        break;
                }
            }
        }

        /// <summary>
        /// iterate each line of logs[] and seperate the data from each line
        /// </summary>
        /// <param name="logs"></param>
        /// <returns>array with only attribute values</returns>
        static private List<Dictionary<string, string>> SeperateAttributes(string[] logs)
        {

            List<Dictionary<string, string>> attributes = new List<Dictionary<string, string>>();

            string jsonString;
            for(int i = 0; i < logs.Length; i++)
            {
                jsonString = logs[i].Split('{', 2)[1]; // seperate elements with "-char and discard every leftout element, with the length of 1 
                Dictionary<string, string> temp = mapElements(jsonString);
                attributes.Add(temp);
            }
            return attributes;
        }
        static private Dictionary<string, string> mapElements(string elements)
        {
            Dictionary<string, string> mappedElements = new Dictionary<string, string>();
            List<string> subs = elements.Split('"').Where(c => c.Length > 1).ToList();
            for (int i = 0; i < subs.Count - 1; i+=2)
            {
                mappedElements.Add(subs[i], subs[i + 1]);
            }
            return mappedElements;
            
        }
        /// <summary>
        /// Read file from a given location and return every line of that file in an array
        /// </summary>
        /// <param name="pathname"> Given pathname</param>
        /// <returns>Array from the rows of the file</returns>
        static private string[] ReadFile(string pathname)
        {
            string extension = Path.GetExtension(pathname);
            string[] file_text;

            if (extension == ".txt")
            {
                file_text = ReadTxt(pathname);
            }
            else file_text = new string[] { extension + " is not included yet" };
            
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