using Microsoft.Office.Interop.Excel;
using _Excel = Microsoft.Office.Interop.Excel;
using Range = Microsoft.Office.Interop.Excel.Range;
using System.Drawing;

/// <summary>
/// Excel class for reading and writing to an excel document
/// @Author Samuli Ryhänen 07.06.2022
/// @Genretech Oy
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
            wb.Close(true);
            excel.Quit();
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
        /// Apply autofilter to the worksheet
        /// </summary>
        public void AddFilter()
        {
            Range range = ws.UsedRange;
            range.Columns.AutoFilter(1, System.Reflection.Missing.Value, XlAutoFilterOperator.xlAnd, System.Reflection.Missing.Value, true);
        }

        /// <summary>
        /// Apply background color  from a cell
        /// </summary>
        /// <param name="row">row index</param>
        /// <param name="column">column index</param>
        /// <param name="color">ARGB value for the color</param>
        public void CellColor(int row, int column, Color color)
        {
            ws.Cells[row, column].Interior.Color = color;
        }

        /// <summary>
        /// change font color from a cell
        /// </summary>
        /// <param name="row">row index</param>
        /// <param name="column">column index</param>
        /// <param name="color">ARGB value for the color</param>
        public void FontColor(int row, int column, Color color)
        {
            ws.Cells[row, column].Font.Color = color;
        }

        /// <summary>
        /// Create a new column for a new occuring attribute
        /// </summary>
        /// <param name="header">new attribute</param>
        /// <param name="value">value of the attribute</param>
        /// <param name="row">row index</param>
        /// <param name="column">column index</param>
        public void WriteNew(string header, string value, int row, int column)
        {
            Write(header, 1, column); // add Header for column
            CellColor(1, column, Color.Black); // column styles
            FontColor(1, column, Color.White); //
            Write(value, row, column); // write data for the cell
        }

        /// <summary>
        /// Write value to a cell
        /// </summary>
        /// <param name="value">value</param>
        /// <param name="row">row index</param>
        /// <param name="column">column index</param>
        public void Write(string value, int row, int column)
        {
            if (row != 1 && column != 2 && row % 2 != 0) {
                ws.Cells[row, column].Interior.Color = Color.LightGray;
            }
            ws.Cells[row, column].Value = value;
        }

        /// <summary>
        /// Wrap cell content into more readable shape
        /// </summary>
        public void FitContent()
        {
            ws.Columns[3].ColumnWidth = 60;
            ws.Columns[3].WrapText = true;
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
