using System.Drawing;

/// <summary>
/// Class for reading UIpath-log messages and inserting them to more readable excel-file
/// @Author Samuli Ryhänen 25.05.2022
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
            AddHeaders(excel);
            List<Dictionary<string, string>> logsList = SeperateAttributes(logs);
            for(int i = 0; i< logsList.Count; i++)
            {
                AddRow(excel, logsList[i], i + 2); // first row is for the headers
            }
            excel.fitContent();
            excel.SaveAs(@"C:\genretech\loginPuhdistus\test\short_test.xlsx");
            excel.Close();
            return "ok";
        }

        /// <summary>
        /// Add headers to first line of excel
        /// </summary>
        /// <param name="excel"></param>
        static private void AddHeaders(Excel excel)
        {
            string[] headers = { "Timestamp", "Log level", "Process name", "Message", "Filename", "Process version", "Robot name", "Machine id", "Fingerprint" };
            for(int i = 0; i < headers.Length; i++)
            {
                excel.Write(headers[i], 1, i+1);
            }

        }

        /// <summary>
        /// Add a single row from the dictionary to excell
        /// </summary>
        /// <param name="excel">excell</param>
        /// <param name="logLine">row added</param>
        /// <param name="row">row index</param>
        static private void AddRow(Excel excel, Dictionary<string, string> logLine, int row)
        {   
            var keys = logLine.Keys;
            
            foreach(string key in keys){
                int column = 10;
                switch (key)
                {
                    case "timeStamp":
                        string time = FormatTime(logLine["timeStamp"]);
                        excel.Write(time, row, 1);
                       break;
                    case "level":
                        excel.Write(logLine["level"], row, 2);
                        Color color = ChooseColor(logLine["level"]);
                        excel.CellColor(row, 2, color);
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
                    case "machineId":
                        excel.Write(logLine["machineId"], row, 8);
                        break;
                    case "jobId":
                        excel.Write(logLine["jobId"], row, 9);
                        break;
                    default:
                        string keyAndVal = key + " : " + logLine[key];
                        excel.Write(keyAndVal, row, column);
                        column++;
                        break;
                }
            }
        }

        static private string FormatTime(string time)
        {
            string formatted = "";
            
            try
            {
                int index = time.Length - 6;
                formatted = time.Replace('T', ' ').Remove(index);
            } //DOTO: Datetime formatting
            catch
            {
                formatted = time;
            }
            return formatted; 
        }

        /// <summary>
        /// Choose a color for a log message level
        /// </summary>
        /// <param name="logLevel">log level</param>
        /// <returns> Color value</returns>
        static private Color ChooseColor(string logLevel)
        {
            Color color;
            switch (logLevel)
            {
                case "Information":
                    color = Color.DarkGreen;
                    break;
                case "Error":
                    color = Color.RosyBrown;
                    break;
                case "Fatal":
                    color = Color.IndianRed;
                    break;
                case "Warning":
                    color = Color.LightGoldenrodYellow;
                    break;
                case "Trace":
                    color = Color.CadetBlue;
                    break;
                case "Verbose":
                    color = Color.DarkGray;
                    break;
                default:
                    color = Color.White;
                    break;
            }
            return color;
        }

        /// <summary>
        /// Create dictionary from each row of logs
        /// </summary>
        /// <param name="logs"></param>
        /// <returns>List dictionary with attributes as key values</returns>
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

        /// <summary>
        /// Modify a row from logs and split in to a dictionary
        /// </summary>
        /// <param name="elements">Row from logs</param>
        /// <returns> dictionary with attribute key and value as value</returns>
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