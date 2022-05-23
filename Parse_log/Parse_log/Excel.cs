using Microsoft.Office.Interop.Excel;
using _Excel = Microsoft.Office.Interop.Excel;
using Range = Microsoft.Office.Interop.Excel.Range;
using System.Drawing;
/// <summary>
/// Excel class for reading and writing to an excel document
/// @author Samuli Ryhänen 23.05.2022
/// </summary>
namespace Parse_log
{
    internal class Excel
    {
        string path = "";
        _Application excel = new _Excel.Application();
        Workbook wb;
        Worksheet ws;

        public Excel(string path, int sheet)
        {
            this.path = path;
            wb = excel.Workbooks.Add();
            ws = excel.Worksheets[sheet];
        }

        public void Close()
        {   
            wb.Close();
        }

        public void SaveAs(string filename)
        {
            wb.SaveAs(filename);
        }

        public void Save()
        {
            wb.Save();
        }

        /// <summary>
        /// Apply background colouyr to a cell
        /// </summary>
        /// <param name="row">row index</param>
        /// <param name="column">column index</param>
        /// <param name="color">ARGB value for the color</param>
        public void CellColor(int row, int column, Color color)
        {
            ws.Cells[row, column].Interior.Color = color;
        }

        /// <summary>
        /// Write value to a cell
        /// </summary>
        /// <param name="value">value</param>
        /// <param name="row">row index</param>
        /// <param name="column">column index</param>
        public void Write(string value, int row, int column)
        {
            ws.Cells[row, column].Value = value;
            
        }
        /// <summary>
        /// autofit content 
        /// </summary>
        public void fitContent()
        {
            Range usedRange = ws.UsedRange;
            usedRange.Columns.AutoFit();
        }

        /// <summary>
        /// Writes to a file in a specific range
        /// </summary>
        /// <param name="starti"></param> starting column index 
        /// <param name="starty"></param> starting row index
        /// <param name="endi"></param> ending column index
        /// <param name="endy"></param> ending row index
        /// <param name="writestring"></param> 2d-array which is written
        public void WriteRange(int starti, int starty, int endi, int endy, string[,] writestring)
        {
            Range range = (Range)ws.Range[ws.Cells[starti, starty], ws.Cells[endi, endy]];
            range.Value2 = writestring;
        }

        /// <summary>
        /// Reads a table in a spesific range
        /// </summary>
        /// <param name="starti"></param> starting column index
        /// <param name="starty"></param> starting row index
        /// <param name="endi"></param> ending column index
        /// <param name="endy"></param> ending row index
        /// <returns> 2d-array </returns>
        public string[,] ReadRange(int starti, int starty, int endi, int endy)
        {
            Range range = (Range)ws.Range[ws.Cells[starti, starty], ws.Cells[endi, endy]];
            object[,] holder = range.Value2;
            string[,] returnstring = new string[endi - starti + 1, endy - starty + 1];
            for (int p = 1; p <= endi - starti + 1; p++)
            {
                for (int q = 1; q <= endy - starty + 1; q++)
                {
                    if (holder[p, q] != null)
                    {
                        returnstring[p - 1, q - 1] = holder[p, q].ToString();
                    }
                }
            }
            return returnstring;
        }
    }
}
