using System.Drawing;

/// <summary>
/// Class for reading UIpath-log messages and inserting them to more readable excel-file
/// @Author Samuli Ryhänen 24.05.2022
/// </summary>
namespace Parse_log
{
    internal class Log
    {
         
        /// <summary>
        /// Process data from a file into an array
        /// </summary>
        /// <param name="pathName"></param>
        public string ProcessData(string pathName)
        {
            string status;
            string file = Path.GetFileNameWithoutExtension(pathName);
            string excelName = file + ".xlsx";
            Excel excel = new Excel(@excelName, 1); // opens first worksheet of excel
            AddHeaders(excel);

            try
            {
                TransferData(pathName, excel);
                status = excelName;
            }
            catch (Exception e)
            {
                status = "Error: " + (e.Message);
            }

            excel.fitContent();
            string path = Path.GetDirectoryName(pathName) + "/" + Path.GetFileNameWithoutExtension(excelName);
            excel.SaveAs(@path);
            excel.Close();
            return status;
        }

        /// <summary>
        /// Add headers to first line of excel
        /// </summary>
        /// <param name="excel"></param>
        static private void AddHeaders(Excel excel)
        {
            string[] headers = { "Timestamp", "Log level", "Process name", "Message", "Filename", "Process version", "Robot name", "Machine id", "Fingerprint" };
            for (int i = 0; i < headers.Length; i++)
            {
                excel.Write(headers[i], 1, i + 1);
            }

        }

        /// <summary>
        /// Read .txt-file from a given path and transfer it directly line by line into an excel file
        /// </summary>
        /// <param name="pathName">path to txt</param>
        /// <param name="excel"> excel </param>
        static void TransferData(string pathName, Excel excel)
        {

            List<Dictionary<string, string>> lines = new List<Dictionary<string, string>>();

            using (StreamReader sr = new StreamReader(pathName))
            {
                String line;
                Dictionary<string, string> mappedLine;
                int i = 2; // first excel row is for headers, so begin from line 2.
                while ((line = sr.ReadLine()) != null)
                {
                    mappedLine = MapElements(line.Split('{', 2)[1]); // need to separate parts before actual data
                    AddExcelRow(excel, mappedLine, i); // insert row directly to an excel
                    i++; // index for excel row
                }
            }
        }

        /// <summary>
        /// Add a single row from the dictionary to excell
        /// </summary>
        /// <param name="excel">excell</param>
        /// <param name="logLine">row added</param>
        /// <param name="row">row index</param>
        static private void AddExcelRow(Excel excel, Dictionary<string, string> logLine, int row)
        {   
            var keys = logLine.Keys;
            int column = 10;
            try
            {
                foreach (string key in keys)
                {

                    switch (key)
                    {
                        case "timeStamp":
                            string time = FormatTime(logLine["timeStamp"]);
                            excel.Write(time, row, 1);
                            //excel.TextFormatOff(1); // remove textformat for timestamp
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
                            string keyAndVal = key + " = " + logLine[key];
                            keyAndVal.Replace(',', ' ').Replace('}', ' ');
                            excel.Write(keyAndVal, row, column);
                            column++;
                            break;
                    }
                }
            }

            catch(Exception e)
            {
                Console.WriteLine("Error inserting excel: " + e);
            }
            
        }

        /// <summary>
        /// Modify a row from logs and split in to a dictionary
        /// </summary>
        /// <param name="elements">Row from logs</param>
        /// <returns> dictionary with attribute key and value as value</returns>
        static private Dictionary<string, string> MapElements(string elements)
        {
            Dictionary<string, string> mappedElements = new Dictionary<string, string>();
            List<string> subs = elements.Split('"').Where(c => c.Length > 1).ToList();
            for (int i = 0; i < subs.Count - 1; i += 2)
            {
                mappedElements.Add(subs[i], subs[i + 1]);
            }
            return mappedElements;

        }

        /// <summary>
        /// Format time to a wanted string 
        /// </summary>
        /// <param name="time">timestamp </param>
        /// <returns>formatted time in a string</returns>
        static private string FormatTime(string time)
        {
            string formatted;
            try
            {
                DateTime datetime = DateTime.ParseExact(time, "yyyy-MM-ddTHH:mm:ss.fffffffzzz",
                                           System.Globalization.CultureInfo.InvariantCulture);
                formatted = datetime.ToString("yyyy-MM-dd HH:mm:ss.fffffff");
            }
            /// seconds have sometimes different amount of decimals
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
    }
}