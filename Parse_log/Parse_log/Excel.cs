using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using Microsoft.Office.Interop.Excel;
using _Excel = Microsoft.Office.Interop.Excel;
using Range = Microsoft.Office.Interop.Excel.Range;

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
        public void CellColor(int row, int column, int colour)
        {
            ws.Cells[row, column].Interior.Color = colour;
        }
        public void Write(string value, int row, int column)
        {
            ws.Cells[row, column].Value = value;
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
