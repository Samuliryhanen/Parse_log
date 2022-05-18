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
        public string processData(string pathname)
        {
            string status;
            try
            {
                string[] logs = readFile(pathname);
                status = addToExcel(logs);
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
        private string addToExcel(string[] logs)
        {
            //Excel excel = new Excel(0); // opens first worksheet of excel
            int row = 0;
            int column = 0;
            string[][] attributes = seperate_attributes(logs);
            for(int i = 0; i< attributes.Length; i++)
            {
                for(int j = 0; j < attributes[i].Length; j++)
                {
                    Console.WriteLine(attributes[i][j]); // TOIMII
                }
            }
            return "ok";
        }

        /// <summary>
        /// TODO: DOKUMENTOI
        /// </summary>
        /// <param name="logs"></param>
        /// <returns></returns>
        private string[][] seperate_attributes(string[] logs)
        {
            string[] subs;
            string[][] attributes = new string[logs.Length][];
            
            for(int i = 0; i < logs.Length; i++)
            {
                subs = logs[i].Split('{', 2);
                attributes[i] = subs[1].Split('"');
            }
            return attributes;
            
        }

        /// <summary>
        /// Read file from a given location and return every line of that file in an array
        /// </summary>
        /// <param name="pathname"> Given pathname</param>
        /// <returns>Array from the rows of the file</returns>
        private string[] readFile(string pathname)
        {
            string extension = Path.GetExtension(pathname);
            string[] file_text = { };

            if (extension == ".txt")
            {
                file_text = readTxt(pathname);
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
        private string[] readTxt(string pathname)
        {
            // Read entire text file content in one string  
            string[] lines = File.ReadAllLines(pathname);
            return lines;
        }
    }
}