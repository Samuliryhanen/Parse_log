using System.Drawing;

/// <summary>
/// Class for reading UIpath-log messages and inserting them to more readable excel-file
/// @Author Samuli Ryhänen 07.06.2022
/// @GenreTech Oy
/// </summary>
namespace Parse_log
{
    //Class to add dynamically new attributes to the excelfile
    public class Headers
    {
        // list for new headers
        private List<string> OverflowingHeaders = new List<string>();

        /// <summary>
        /// Add new member to OverflowingHeaders
        /// </summary>
        /// <param name="header"></param>
        /// <returns></returns>
        public int AddHeader(string header)
        {
            OverflowingHeaders.Add(header);
            return OverflowingHeaders.Count;
        }
        /// <summary>
        /// Find index of an element from OverflowingHeaders
        /// </summary>
        /// <param name="element"></param>
        /// <returns></returns>
        public int FindElementIndex(string element)
        {
            int index = OverflowingHeaders.IndexOf(element);
            if (index != -1)
            {
                index++;
            }
            return index; 
        }
        /// <summary>
        /// Get OverflowingHeaders
        /// </summary>
        /// <returns>List<String, String>OverflowingHeaders </returns>
        public List<string> GetValues()
        {
            return OverflowingHeaders;
        }
        /// <summary>
        /// Get count of OverflowingHeaders
        /// </summary>
        /// <returns>integer</returns>
        public int GetCount()
        {
            return OverflowingHeaders.Count;
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
        public void ProcessData(string pathNameIn, string pathNameOut)
        {
            try
            {
                Excel excel = new Excel(pathNameOut, 1); // opens new excel, sheet no.1
                Headers headers = new Headers();
                try
                {
                    TransferData(pathNameIn, excel, headers);
                    excel.AddFilter();
                    excel.FitContent();
                    excel.SaveAs(pathNameOut);
                    excel.Close();
                }
                catch (Exception e)
                {
                    Console.WriteLine("Error transfering data " + e);
                    excel.Close();
                }
            }
            catch(Exception e)
            {
                Console.WriteLine("Error opening excel " + e);
            }

        }

        /// <summary>
        /// Read from a given path and transfer it directly line by line into an excel file
        /// </summary>
        /// <param name="pathName">path to txt</param>
        /// <param name="excel"> excel </param>
        static void TransferData(string pathName, Excel excel, Headers headers)
        {
            AddHeader("timestamp",excel, 1, headers);
            AddHeader("level", excel, 2, headers);
            
            using (StreamReader sr = new StreamReader(pathName))
            {
                string line;
                Dictionary<string, string> mappedLine;
                int i = 2; // first excel row is for headers, so begin from line 2.
                while ((line = sr.ReadLine()) != null)
                {
                    try
                    {
                        mappedLine = MapElements(line); // need to separate parts before actual data
                        AddExcelRow(excel, mappedLine, i, headers); // insert row directly to an excel
                        i++;// index for excel row
                    }
                    catch
                    {
                        Console.WriteLine("Error inserting rows!");
                        excel.Close();
                    }
                }
            }
        }

        /// <summary>
        /// Add a single row from the dictionary to excell
        /// </summary>
        /// <param name="excel">excel</param>
        /// <param name="logLine">row added</param>
        /// <param name="row">row index for the excel</param>
        static private void AddExcelRow(Excel excel, Dictionary<string, string> logLine, int row, Headers headers)
        {   
            var keys = logLine.Keys;
            try
            {
                foreach (string key in keys)
                {
                    int columnIndex = headers.FindElementIndex(key);

                    switch (columnIndex)
                    {
                        case 1: // timestamp wanted to the first cell
                            string time = FormatTime(logLine[key]);
                            excel.Write(time, row, 1);
                            break;
                        case 2: // second cell is for log-level and color code
                            excel.Write(logLine[key], row, 2);
                            Color color = ChooseColor(logLine[key]);
                            if(logLine[key].ToLower() != "warning") // better contrast 
                            {
                                excel.FontColor(row, 2, Color.White);
                            }
                            
                            excel.CellColor(row, 2, color);
                            break;
                        default:
                            if (columnIndex == -1)
                            {
                                columnIndex = headers.GetCount() + 1;       
                                AddHeader(key, excel, columnIndex, headers);// If a new attribute add to a new column attribute name 
                                excel.Write(logLine[key], row, columnIndex);// and value from the attribute
                            }
                            else
                            {
                                excel.Write(logLine[key], row, columnIndex); // add value if attribute already exists
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
        /// Add new header to excel file
        /// </summary>
        /// <param name="name"></param>
        /// <param name="excel"></param>
        /// <param name="index"></param>
        static private void AddHeader(string name, Excel excel, int index, Headers headers)
        {
            try
            {
                excel.Write(name, 1, index);
                excel.CellColor(1, index, Color.Black);
                excel.FontColor(1, index, Color.White);
                headers.AddHeader(name);
            }
            catch
            {
                Console.WriteLine("Error adding headers! ");
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
            elements = elements.Split('"', 2)[1]; // seperate from beginning of the log text
            Dictionary<string, string> mappedElements = new Dictionary<string, string>();
            string[] strings = elements.Split(','); // seperate each individual element from each other
            string[] temp;
            string prevKey="";
            foreach (string j in strings)
            {
                try
                {
                    temp = j.Replace('"', ' ').Trim().Split(":", 2); //remove quotations and seperate each attribute value from its name
                    string key = temp[0].ToLower().Trim();
                    
                    if (temp.Length == 1) // log-file can include Chars such as ":" or "," , which breaks the sorting
                    {
                        mappedElements[prevKey] = mappedElements[prevKey] + " " + temp[0].Trim();
                        continue;
                    }
                    if (mappedElements.ContainsKey(key)) // if broken data contains existing key-name, just add the value to a proper cell
                    {
                        mappedElements[key] = mappedElements[key] + " " + temp[1].Trim();
                        continue;
                    }
                    prevKey = key;
                    mappedElements.Add(key, temp[1].Trim());

                }
                catch(Exception e)
                {
                    Console.WriteLine(e);
                }
            }
            return mappedElements;
        }

        /// <summary>
        /// Format time 
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
                formatted = datetime.ToString("@yyyy-MM-dd HH:mm:ss.fffffff"); // excel will omit autoformat with @- char 
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
        /// contrast checked with https://webaim.org/resources/contrastchecker/
        /// </summary>
        /// <param name="logLevel">log level</param>
        /// <returns> Color value</returns>
        static private Color ChooseColor(string logLevel)
        {
            Color color;
            switch (logLevel)
            {
                case "Information":
                    color = Color.FromArgb(44, 96, 1);
                    break;
                case "Error":
                    color = Color.FromArgb(163, 49, 0);
                    break;
                case "Fatal":
                    color = Color.FromArgb(179, 0, 0); 
                    break;
                case "Warning":
                    color = Color.LightGoldenrodYellow; // okay as it is
                    break;
                case "Trace":
                    color = Color.FromArgb(37, 72, 212);
                    break;
                case "Verbose":
                    color = Color.Black;
                    break;

                default:
                    color = Color.White;
                    break;
            }
            return color;
        }
    }
}