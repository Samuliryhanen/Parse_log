using System.Drawing;
using System.Text.RegularExpressions;

/// <summary>
/// Class for reading UIpath-log messages and inserting them to more readable excel-file
/// @Author Samuli Ryhänen 24.05.2022
/// </summary>
namespace Parse_log
{
    //class to handle dilemma with overflown attributes
    public class Headers
    {
        private List<string> overflowingHeaders = new List<string>(); // list for headers, which are not defined in switch-case sorting

        public int AddHeader(string header)
        {
            overflowingHeaders.Add(header);
            return overflowingHeaders.Count;
        }
        public int FindElementIndex(string element)
        {
            return overflowingHeaders.IndexOf(element);
        }
        public List<string> GetValues()
        {
            return overflowingHeaders;
        }
        public int GetLength()
        {
            return overflowingHeaders.Count;
        }
    }

    /// <summary>
    /// Actual log-class
    /// </summary>
    internal class Log
    {
        /// <summary>
        /// Process data from a file into an array
        /// </summary>
        /// <param name="pathName"></param>
        public string ProcessData(string pathName)
        {
            string status;
            string file;
            string excelName;
            try
            {
                file = Path.GetFileNameWithoutExtension(pathName);
                excelName = file + ".xlsx";
                Excel excel = new Excel(@excelName, 1); // opens first worksheet of excel
                Headers headers = new Headers();
                try
                {
                    TransferData(pathName, excel, headers);
                    status = excelName;
                    excel.FitContent();
                    string path = Path.GetDirectoryName(pathName)+ @"\" + Path.GetFileNameWithoutExtension(excelName);
                    excel.SaveAs(@path);
                    excel.Close();
                }
                catch (Exception e)
                {
                    status = "Error: " + (e.Message);
                    excel.Close();
                }
                return status;
            }
            catch(Exception e)
            {
                return "Error: " + e;
            }

        }

        /// <summary>
        /// Read .txt-file from a given path and transfer it directly line by line into an excel file
        /// </summary>
        /// <param name="pathName">path to txt</param>
        /// <param name="excel"> excel </param>
        static void TransferData(string pathName, Excel excel, Headers headers)
        {
            AddHeaders(excel, headers);
            List<Dictionary<string, string>> lines = new List<Dictionary<string, string>>();

            using (StreamReader sr = new StreamReader(pathName))
            {
                String line;
                Dictionary<string, string> mappedLine;
                int i = 2; // first excel row is for headers, so begin from line 2.
                while ((line = sr.ReadLine()) != null)
                {
                    mappedLine = MapElements(line.Split('{', 2)[1]); // need to separate parts before actual data
                    AddExcelRow(excel, mappedLine, i, headers); // insert row directly to an excel
                    i++; // index for excel row
                }
                try
                {
                    int columnCount = headers.GetLength() + 9;
                    for (int j = 1; j <= columnCount; j++)
                    {
                        excel.AddFilter(1, j);
                    }
                }
                catch
                {
                    Console.WriteLine("Error while inserting filters");
                    excel.Close();
                }
                
            }
        }

        /// <summary>
        /// Add headers to first line of excel
        /// </summary>
        /// <param name="excel"></param>
        static private void AddHeaders(Excel excel, Headers headers)
        {
            try
            {
                string[] values = { "Timestamp", "Log level", "Process name", "Message", "Filename", "Process version", "Robot name", "Machine id", "Fingerprint" };
                for (int i = 0; i < values.Length; i++)
                {
                    excel.Write(values[i], 1, i + 1);
                    excel.CellColor(1, i + 1, Color.Black);
                    excel.FontColor(1, i + 1, Color.White);
                }
            }
            catch
            {
                Console.WriteLine("Error adding headers! ");
                excel.Close();
            }
        }


        /// <summary>
        /// Add a single row from the dictionary to excell
        /// </summary>
        /// <param name="excel">excell</param>
        /// <param name="logLine">row added</param>
        /// <param name="row">row index</param>
        static private void AddExcelRow(Excel excel, Dictionary<string, string> logLine, int row, Headers headers)
        {   
            var keys = logLine.Keys;
            try
            {
                foreach (string key in keys)
                {
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
                            int index = headers.FindElementIndex(key); // header list is shorter than all the headers in excel file
                            if (index != -1)
                            {
                                index += 9;
                                excel.Write(logLine[key], row, index);
                            }
                            else
                            {
                                index = 9 + headers.AddHeader(key); // header list is shorter than all the headers in excel file
                                excel.WriteNew(key, logLine[key], row, index);
                            }
                            break;
                    }
                }
            }

            catch(Exception e)
            {
                Console.WriteLine("Error inserting excel: " + e);
                excel.Close();
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
            List<string> subs = elements.Split('"').Where(c => c.Length > 1 || (c.Length == 1 && !Char.IsPunctuation(char.Parse(c)))).ToList();
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