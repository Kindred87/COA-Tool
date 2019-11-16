using System;
using System.Collections.Generic;
using System.Text;
using OfficeOpenXml;
using System.Xml;
using System.Drawing;
using OfficeOpenXml.Style;
using System.IO;
using OfficeOpenXml.Drawing;
using CoA_Tool.Utility;

namespace CoA_Tool.Excel
{
    /// <summary>
    /// Handles processes relating to a finished good datasheet
    /// </summary>
    class FinishedGoodsData
    {
        
        /// <summary>
        /// File contents arranged as a pseudo-grid, where [y][x], see Load() for further info
        /// </summary>
        public List<List<string>> Contents; // TODO: Refactor to dictionaries
        
        public FinishedGoodsData()
        {
        }
        /// <summary>
        /// Populates Contents with data from finished goods Excel document
        /// </summary>
        /// <param name="filePath"></param>
        /// <returns></returns>
        public void Load()
        {
            string filePath = GetFilePath();

            FileInfo excelFile = new FileInfo(filePath);

            Contents = new List<List<string>>();

            using (ExcelPackage package = new ExcelPackage(excelFile))
            {
                Console.Util.WriteMessageInCenter("Loading finished goods data...");

                for (int i = 2; i < package.Workbook.Worksheets[1].Dimension.Rows; i++)
                {
                    Contents.Add(new List<string>());

                    Contents[i - 2].Add(package.Workbook.Worksheets[1].Cells[i, 1].Value.ToString()); // Part codes
                    Contents[i - 2].Add(package.Workbook.Worksheets[1].Cells[i, 2].Value.ToString()); // Part description
                    Contents[i - 2].Add(package.Workbook.Worksheets[1].Cells[i, 7].Value.ToString()); // Days to expiry
                    Contents[i - 2].Add(package.Workbook.Worksheets[1].Cells[i, 8].Value.ToString()); // Recipe
                }

                Console.Util.RemoveMessageInCenter();
            }
        }
        /// <summary>
        /// Gets path of Excel file, searching in the Desktop and Downloads folders.
        /// </summary>
        /// <returns></returns>
        private string GetFilePath()
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
                return GetFilePath(); // Until the file is found
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
        /// <summary>
        /// Retrieves the name associated with the provided product code
        /// </summary>
        /// <param name="productCode">The five digit finished good product code</param>
        /// <returns></returns>
        public bool ProductNameExists(string productCode, out string productName)
        {
            foreach (List<string> line in Contents) 
            {
                if (line[0] == productCode)
                {
                    productName = line[1];
                    return true;
                }
            }
            productName = string.Empty;
            return false;
        }
        /// <summary>
        /// Retrieves the recipe code associated with the provided product code
        /// </summary>
        /// <param name="productCode">The five digit finished good product code</param>
        /// <returns></returns>
        public bool RecipeCodeExists(string productCode, out string recipeCode)
        {
            recipeCode = "";

            foreach (List<string> line in Contents) 
            {
                if (line[0] == productCode)
                {
                    recipeCode = line[3];
                    return true;
               }
            }
            
            return false;
        }
        /// <summary>
        /// Determines the made date of a product based on its lot code
        /// </summary>
        /// <param name="productCode">The product code for the product in question</param>
        /// <param name="lotCode">The lotcode for the product in question</param>
        /// <returns></returns>
        public DateTime GetMadeDate(string lotCode)
        {
            string productCode = Lot.ProductCode(lotCode);
            int daysToExpiry = 0;

            foreach (List<string> line in Contents) 
            {
                if (line[0] == productCode)
                {
                    daysToExpiry = Convert.ToInt32(line[2]);
                }
            }

            int expiryDay = Convert.ToInt32(lotCode[9].ToString()) * 10;
            expiryDay += Convert.ToInt32(lotCode[10].ToString());

            int expiryMonth = Convert.ToInt32(lotCode[7].ToString()) * 10;
            expiryMonth += Convert.ToInt32(lotCode[8].ToString());

            int expiryYear = Convert.ToInt32(lotCode[11].ToString()) * 10;
            expiryYear += Convert.ToInt32(lotCode[12].ToString());
            expiryYear += 2000;

            return new DateTime(expiryYear, expiryMonth, expiryDay).AddDays(daysToExpiry * -1);
        }
    }
}
