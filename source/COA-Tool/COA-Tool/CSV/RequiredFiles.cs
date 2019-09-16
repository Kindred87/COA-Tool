using System;
using System.Collections.Generic;
using System.Text;
using System.IO;
using System.Linq;
using System.Xml;
using System.Drawing;
using OfficeOpenXml.Style;
using OfficeOpenXml.Drawing;
using OfficeOpenXml;

namespace COA_Tool.CSV
{

    class RequiredFiles
    {
        public List<List<string>> DelimitedMicroResults = new List<List<string>>();
        public List<List<string>> DelimitedTitrationResults = new List<List<string>>();
        // List<string = {Part code, part description, days to expiry, recipe}
        public List<List<string>> FinishedGoods = new List<List<string>>();
        // List<string> = {Recipe code, days to expiry}
        public List<List<string>> Recipes = new List<List<string>>();
        public bool AllFilesReady;
        public RequiredFiles()
        {
            List<string> filePaths = GetFileNamesFromDesktop();
            List<List<string>> csvFileContents = LoadCSVFiles(filePaths);


            if (FilesLoaded(csvFileContents.AsReadOnly()) == true)
            {

                AllFilesReady = true;

                foreach (string line in csvFileContents[IndexOfMicroResults(csvFileContents)])
                {
                    DelimitedMicroResults.Add(line.Split(new char[] { ',' }).ToList());
                }
                foreach (string line in csvFileContents[IndexOfTitrationResults(csvFileContents)])
                {
                    DelimitedTitrationResults.Add(line.Split(new char[] { ',' }).ToList());
                }

                LoadExcelFiles(filePaths);
            }
            else
                AllFilesReady = false;
        }
        /// <summary>
        /// Checks for files based on criteria
        /// </summary>
        /// <param name="fileContents"></param>
        public bool FilesLoaded(IReadOnlyList<List<string>> fileContents)
        {
            if (CheckForMicroResults(fileContents) == true)
            {
                System.Console.SetCursorPosition(10, 0);
                System.Console.Write("Micro CSV: ");
                System.Console.ForegroundColor = ConsoleColor.Green;
                System.Console.Write("Loaded");
                System.Console.ForegroundColor = ConsoleColor.Gray;
            }
            else
            {
                System.Console.SetCursorPosition(10, 0);
                System.Console.Write("Micro CSV: ");
                System.Console.ForegroundColor = ConsoleColor.Red;
                System.Console.Write("Missing");
                System.Console.ForegroundColor = ConsoleColor.Gray;
                return false;
            }

            if (CheckForTitrationResults(fileContents) == true)
            {
                System.Console.SetCursorPosition(7, 1);
                System.Console.Write("Dressing CSV: ");
                System.Console.ForegroundColor = ConsoleColor.Green;
                System.Console.Write("Loaded");
                System.Console.ForegroundColor = ConsoleColor.Gray;
            }
            else
            {
                System.Console.SetCursorPosition(7, 1);
                System.Console.Write("Dressing CSV: ");
                System.Console.ForegroundColor = ConsoleColor.Red;
                System.Console.Write("Missing");
                System.Console.ForegroundColor = ConsoleColor.Gray;
                return false;
            }

            return true;

        }
        /// <summary>
        /// Gets the file paths of .csv files from the user's desktop directory
        /// </summary>
        /// <returns></returns>
        private List<string> GetFileNamesFromDesktop()
        {
            Console.Util.WriteMessageInCenter("Locating files...");

            string desktopPath = Environment.GetFolderPath(Environment.SpecialFolder.Desktop);

            List<string> fileNames = Directory.GetFiles(desktopPath, "*.csv").ToList();
            foreach(string additionalPath in Directory.GetFiles(desktopPath, "*.xlsx"))
            {
                fileNames.Add(additionalPath);
            }

            return fileNames;
        }
        /// <summary>
        /// Reads files to a List<string>
        /// </summary>
        /// <param name="filePaths"></param>
        /// <returns></returns>
        private List<List<string>> LoadCSVFiles(List<string> filePaths)
        {
            Console.Util.WriteMessageInCenter("Loading CSV files...");

            List<List<string>> csvFileContents = new List<List<string>>();

            foreach (string path in filePaths)
            {
                if(path.EndsWith(".csv") || path.EndsWith(".CSV"))
                    csvFileContents.Add(File.ReadAllLines(path).ToList());
                // TODO: Filter csv files before reading lines
            }

            Console.Util.RemoveMessageInCenter();
            return csvFileContents;
        }
        private void LoadExcelFiles(List<string> filePaths)
        {
            Console.Util.WriteMessageInCenter("Loading Excel files...");

            foreach(string path in filePaths)
            {
                if(path.EndsWith(".xlsx") || path.EndsWith(".XLSX"))
                {
                    FileInfo excelFile = new FileInfo(path);

                    using (ExcelPackage package = new ExcelPackage(excelFile))
                    {
                        // There are 6 worksheets due to a hidden Vlookup sheet
                        if(package.Workbook.Worksheets.Count == 6)
                        {
                            if(package.Workbook.Worksheets[1].Name == "FG 1 NO. + 59 NO.")
                            {
                                //FinishedGoods = PopulateFinishedGoods(package.Workbook.Worksheets[1]);
                                PopulateFinishedGoods(package.Workbook.Worksheets[1]);
                            }

                            if(package.Workbook.Worksheets[6].Name == "All Recipes 5 NO.")
                            {
                                PopulateRecipes(package.Workbook.Worksheets[6]);
                            }
                        }
                       
                    }
                }
            }

            Console.Util.RemoveMessageInCenter();

        }
        /// <summary>
        /// Intended for finished goods data, fetches certain values from each row in the provided worksheet
        /// </summary>
        /// <param name="worksheet"></param>
        private void PopulateFinishedGoods(ExcelWorksheet worksheet)
        {
            // Skips header row
            for (int i = 2; i <= worksheet.Dimension.Rows; i++)
            {
                FinishedGoods.Add(new List<string>());

                FinishedGoods[i - 2].Add(worksheet.Cells[i, 1].Value.ToString());
                FinishedGoods[i - 2].Add(worksheet.Cells[i, 2].Value.ToString());
                FinishedGoods[i - 2].Add(worksheet.Cells[i, 7].Value.ToString());
                FinishedGoods[i - 2].Add(worksheet.Cells[i, 8].Value.ToString());
            }
        }
        /// <summary>
        /// Intended for recipe data, fetches certain values from each row in the provided worksheet
        /// </summary>
        /// <param name="worksheet"></param>
        /// <returns></returns>
        private void PopulateRecipes(ExcelWorksheet worksheet)
        {
            // Skips header row
            for(int i = 2; i <= worksheet.Dimension.Rows; i++)
            {
                Recipes.Add(new List<string>());

                Recipes[i - 2].Add(worksheet.Cells[i, 1].Value.ToString());
                Recipes[i - 2].Add(worksheet.Cells[i, 7].Value.ToString());
            }
        }
        /// <summary>
        /// Deletes all files from given array of paths
        /// </summary>
        /// <param name="filePaths"></param>
        private void DeleteFiles(string[] filePaths)
        {
            foreach (string path in filePaths)
            {
                File.Delete(path);
            }
        }
        /// <summary>
        /// Checks whether one of the given lists contains micro results
        /// </summary>
        /// <param name="fileContents"></param>
        /// <returns></returns>
        private bool CheckForMicroResults(IReadOnlyList<List<string>> fileContents)
        {
            char[] delimiter = { ',' };

            for (int i = 0; i < fileContents.Count; i++)
            {
                string extract = fileContents[i][0].Split(delimiter, StringSplitOptions.RemoveEmptyEntries)[0];

                if (extract == "H1" || extract == "L1" || extract == "S1")
                {
                    return true;
                }
            }
            return false;
        }
        /// <summary>
        /// Determines index of list containing micro results, returns -1 if no list is found
        /// </summary>
        /// <param name="fileContents"></param>
        /// <returns></returns>
        private int IndexOfMicroResults(IReadOnlyList<List<string>> fileContents)
        {
            char[] delimiter = { ',' };

            for (int i = 0; i < fileContents.Count; i++)
            {
                string extract = fileContents[i][0].Split(delimiter, StringSplitOptions.RemoveEmptyEntries)[0];

                if (extract == "H1" || extract == "L1" || extract == "S1")
                {
                    return i;
                }
            }
            return -1;
        }
        /// <summary>
        /// Checks whether one of the given lists contains titration results
        /// </summary>
        /// <param name="fileContents"></param>
        /// <returns></returns>
        private bool CheckForTitrationResults(IReadOnlyList<List<string>> fileContents)
        {
            char[] delimiter = { ',' };

            for (int i = 0; i < fileContents.Count; i++)
            {
                string extract = fileContents[i][0].Split(delimiter, StringSplitOptions.RemoveEmptyEntries)[6];

                if (extract == "Hurricane" || extract == "Lowell" || extract == "Sandpoint")
                {
                    return true;
                }
            }
            return false;
        }
        /// <summary>
        /// Determines index of list containing titration results, returns -1 if no list is found
        /// </summary>
        /// <param name="fileContents"></param>
        /// <returns></returns>
        private int IndexOfTitrationResults(IReadOnlyList<List<string>> fileContents)
        {
            char[] delimiter = { ',' };

            for (int i = 0; i < fileContents.Count; i++)
            {
                string extract = fileContents[i][0].Split(delimiter, StringSplitOptions.RemoveEmptyEntries)[6];

                if (extract == "Sandpoint" || extract == "Hurricane" || extract == "Lowell")
                {
                    return i;
                }
            }
            return -1;
        }
    }
}
