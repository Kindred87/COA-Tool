using System;
using System.Collections.Generic;
using System.Text;
using OfficeOpenXml;
using System.Xml;
using System.Drawing;
using OfficeOpenXml.Style;
using System.IO;
using OfficeOpenXml.Drawing;

namespace CoA_Tool.Excel
{
    
    class FinishedGoods
    {
        // File contents arranged as a grid.  See FileContents() for value descriptions.
        public List<List<string>> Contents;
        public FinishedGoods()
        {
            string path = FilePath();
            Contents = FileContents(path);
        }
        /// <summary>
        /// Extracts needed values from the first worksheet of the provided Excel file
        /// </summary>
        /// <param name="filePath"></param>
        /// <returns></returns>
        private List<List<string>> FileContents(string filePath)
        {
            List<List<string>> contents = new List<List<string>>();

            FileInfo excelFile = new FileInfo(filePath);

            using (ExcelPackage package = new ExcelPackage(excelFile))
            {
                Console.Util.WriteMessageInCenter("Loading finished goods data...");

                for (int i = 2; i < package.Workbook.Worksheets[1].Dimension.Rows; i++)
                {
                    contents.Add(new List<string>());

                    contents[i - 2].Add(package.Workbook.Worksheets[1].Cells[i, 1].Value.ToString()); // Part codes
                    contents[i - 2].Add(package.Workbook.Worksheets[1].Cells[i, 2].Value.ToString()); // Part description
                    contents[i - 2].Add(package.Workbook.Worksheets[1].Cells[i, 7].Value.ToString()); // Days to expiry
                    contents[i - 2].Add(package.Workbook.Worksheets[1].Cells[i, 8].Value.ToString()); // Recipe
                }
            }

            return contents;
        }
        /// <summary>
        /// Gets path of Excel file, searching in the Desktop and Downloads folders.
        /// </summary>
        /// <returns></returns>
        private string FilePath()
        {
            string desktopPath = Environment.GetFolderPath(Environment.SpecialFolder.Desktop);
            string downloadsPath = Environment.GetFolderPath(Environment.SpecialFolder.UserProfile);
            downloadsPath += "\\Downloads";

            Console.Util.WriteMessageInCenter("Locating finished goods file...");

            string path; // Did not inline instantiation due to clarity concerns

            if (SearchDirectory(desktopPath, out path) ||  SearchDirectory(downloadsPath, out path))
            {
                Console.Util.RemoveMessageInCenter();
                return path;
            }
            else
            {
                Console.Util.WriteMessageInCenter("Finished goods file could not be located.  Press a key to search again.", ConsoleColor.Red);
                System.Console.ReadKey();
                return FilePath(); // Until the file is found
            }
        }
        /// <summary>
        /// Searches directory, and all sub-directories, for Excel file with certain cell values
        /// </summary>
        /// <param name="dir"></param>
        /// <param name="filePath"></param>
        /// <returns></returns>
        private bool SearchDirectory(string dir, out string filePath)
        {
            foreach(string path in Directory.EnumerateFiles(dir, "*.*", SearchOption.AllDirectories))
            {
                if (IsFile(path))
                {
                    filePath = path;
                    return true;
                }
            }

            filePath = string.Empty;
            return false;
        }
        /// <summary>
        /// Checks if file is an Excel spreadsheet and identifies target via cell values
        /// </summary>
        /// <param name="path"></param>
        /// <returns></returns>
        private bool IsFile(string path)
        {
            if (path.EndsWith(".xlsx") || path.EndsWith(".XLSX"))
            {
                FileInfo excelFile = new FileInfo(path);

                using (ExcelPackage package = new ExcelPackage(excelFile))
                {
                    foreach (ExcelWorksheet ws in package.Workbook.Worksheets)
                    {
                        // Currently the most reliable criteria
                        if (ws.Name == "FG 1 NO. + 59 NO.")
                            return true;
                    }
                }
                return false;
            }
            else
                return false;
        }
    }
}
