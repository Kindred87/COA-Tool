using System;
using System.Collections.Generic;
using System.Text;
using System.IO;
using System.Linq;

namespace COA_Tool.CSV
{

    class RequiredFiles
    {
        public List<List<string>> DelimitedFinishedGoods = new List<List<string>>();
        public List<List<string>> DelimitedMicroResults = new List<List<string>>();
        public List<List<string>> DelimitedTitrationResults = new List<List<string>>();
        public RequiredFiles()
        {
            string[] filePaths = GetFileNamesFromDesktop();
            List<List<string>> fileContents = LoadFiles(filePaths);

            if (FilesLoaded(fileContents.AsReadOnly()) == true)
            {
                foreach (string line in fileContents[IndexOfFinishedGoods(fileContents)])
                {
                    DelimitedFinishedGoods.Add(line.Split(new char[] { ',' }).ToList());
                }
                foreach (string line in fileContents[IndexOfMicroResults(fileContents)])
                {
                    DelimitedMicroResults.Add(line.Split(new char[] { ',' }).ToList());
                }
                foreach (string line in fileContents[IndexOfTitrationResults(fileContents)])
                {
                    DelimitedTitrationResults.Add(line.Split(new char[] { ',' }).ToList());
                }

            }
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

            if (CheckForFinishedGoods(fileContents) == true)
            {
                System.Console.SetCursorPosition(1, 2);
                System.Console.Write("Finished Goods CSV: ");
                System.Console.ForegroundColor = ConsoleColor.Green;
                System.Console.Write("Loaded");
                System.Console.ForegroundColor = ConsoleColor.Gray;
            }
            else
            {
                System.Console.SetCursorPosition(1, 2);
                System.Console.Write("Finished Goods CSV: ");
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
        private string[] GetFileNamesFromDesktop()
        {
            Console.Util.WriteMessageInCenter("Locating files...");

            string desktopPath = Environment.GetFolderPath(Environment.SpecialFolder.Desktop);

            string[] fileNames = Directory.GetFiles(desktopPath, "*.csv");

            return fileNames;
        }
        /// <summary>
        /// Reads files to a List<string>
        /// </summary>
        /// <param name="filePaths"></param>
        /// <returns></returns>
        private List<List<string>> LoadFiles(string[] filePaths)
        {
            Console.Util.WriteMessageInCenter("Loading files...");

            List<List<string>> csvFileContents = new List<List<string>>();

            foreach (string path in filePaths)
            {
                csvFileContents.Add(File.ReadAllLines(path).ToList());
            }

            return csvFileContents;
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
        /// <summary>
        /// Checks whether one of the given lists contains finished good information
        /// </summary>
        /// <param name="fileContents"></param>
        /// <returns></returns>
        private bool CheckForFinishedGoods(IReadOnlyList<List<string>> fileContents)
        {
            char[] delimiter = { ',' };

            for (int i = 0; i < fileContents.Count; i++)
            {
                string extract = fileContents[i][0].Split(delimiter, StringSplitOptions.RemoveEmptyEntries)[0];

                if (extract == "Part Code")
                {
                    return true;
                }
            }
            return false;
        }
        /// <summary>
        /// Determines index of list containing finished good information, returns -1 if no list is found
        /// </summary>
        /// <param name="fileContents"></param>
        /// <returns></returns>
        private int IndexOfFinishedGoods(IReadOnlyList<List<string>> fileContents)
        {
            char[] delimiter = { ',' };

            for (int i = 0; i < fileContents.Count; i++)
            {
                string extract = fileContents[i][0].Split(delimiter, StringSplitOptions.RemoveEmptyEntries)[0];

                if (extract == "Part Code")
                {
                    return i;
                }
            }
            return -1;
        }
    }
}
