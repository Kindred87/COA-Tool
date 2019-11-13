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

namespace CoA_Tool.CSV
{
    /// <summary>
    /// Handles location and loading of CSV dumps from NWA
    /// </summary>
    class NWAData
    {
        // Class variables
        //  Lists
        public List<List<string>> DelimitedMicroResults = new List<List<string>>();
        public List<List<string>> DelimitedTitrationResults = new List<List<string>>();
        public enum TitrationOffset { Acidity = 5, Viscosity = 2, Salt = 4, pH = 6 }
        public enum MicroOffset { Yeast = 9, Mold = 11, Aerobic = 15, Coliform = 7, Lactic = 13, EColiform = 5 }

        // Constructor
        public NWAData()
        {

        }

        // Public methods
        /// <summary>
        /// Finds needed files and populates class lists with their contents
        /// </summary>
        /// <param name="filePaths"></param>
        /// <returns></returns>
        public void LoadCSVFiles()
        {
            string microPath = string.Empty;
            string titrationPath = string.Empty;

            while (microPath == string.Empty || titrationPath == string.Empty)
            {
                GetFilePaths(out microPath, out titrationPath);

                if(microPath == string.Empty)
                {
                    Console.Util.WriteMessageInCenter("Could not locate micro data on the desktop." +  
                        "  Press a key once the file is ready to be loaded.", ConsoleColor.Red);
                    System.Console.ReadKey();
                    System.Console.Write("\b \b ");
                }
                if(titrationPath == string.Empty)
                {
                    Console.Util.WriteMessageInCenter("Could not locate titration data on the desktop." + 
                        "  Press a key once the file is ready to be loaded.", ConsoleColor.Red);
                    System.Console.ReadKey();
                    System.Console.Write("\b \b ");
                }
            }

            if (ReloadFileBecauseTooOld(microPath, titrationPath))
                LoadCSVFiles();

            Console.Util.WriteMessageInCenter("Loading micro data...");
            foreach (string line in File.ReadLines(microPath))
            {
                DelimitedMicroResults.Add(line.Split(new char[] { ',' }).ToList());
            }
            Console.Util.WriteMessageInCenter("Loading titration data...");
            foreach (string line in File.ReadLines(titrationPath))
            {
                DelimitedTitrationResults.Add(line.Split(new char[] { ',' }).ToList());
            }

            Console.Util.RemoveMessageInCenter();
        }

        // Private methods
        /// <summary>
        /// Prompts user to optionally reload files if last update meets or exceeds 6 hours
        /// </summary>
        /// <param name="microPath"></param>
        /// <param name="titrationPath"></param>
        /// <returns></returns>
        private bool ReloadFileBecauseTooOld(string microPath, string titrationPath)
        {
            int hoursSinceLastMicroUpdate = (int)(DateTime.Now - File.GetLastWriteTime(microPath)).TotalHours;
            int hoursSinceLastTitrationUpdate = (int)(DateTime.Now - File.GetLastWriteTime(titrationPath)).TotalHours;


            if (hoursSinceLastMicroUpdate >= 6 || hoursSinceLastTitrationUpdate >= 6)
            {
                string choicePrompt;

                if (hoursSinceLastMicroUpdate >= 6)
                {
                    choicePrompt = "Micro data hasn't been updated in " + hoursSinceLastMicroUpdate;
                }
                else
                {
                    choicePrompt = "Titration data hasn't been updated in " + hoursSinceLastTitrationUpdate;
                }

                choicePrompt += " hours.  Would you first like to update the file and reload it? Press Y or N";

                Console.Util.WriteMessageInCenter(choicePrompt);

                if (System.Console.ReadKey().Key == ConsoleKey.Y)
                {
                    System.Console.Write("\b \b ");
                    Console.Util.WriteMessageInCenter("Press any key once the file is ready for reloading.");
                    System.Console.ReadKey();
                    System.Console.Write("\b \b ");
                    Console.Util.RemoveMessageInCenter();
                    return true;
                }
                else
                {
                    System.Console.Write("\b \b "); // Deletes key char entered in if condition
                    Console.Util.RemoveMessageInCenter();
                    return false;
                }
            }
            else
                return false;
        }
        /// <summary>
        /// Searches the desktop for needed CSV files 
        /// </summary>
        /// <param name="files"></param>
        /// <returns></returns>
        private void GetFilePaths(out string MicroFilePath, out string TitrationFilePath)
        {
            MicroFilePath = string.Empty;
            TitrationFilePath = string.Empty;
            string fileName;
            char[] delimiter = { '/', '\\', };
            

            foreach (string csvFile in Directory.GetFiles(Environment.GetFolderPath(Environment.SpecialFolder.Desktop), "*.csv"))
            {
                string firstLine = string.Empty;

                fileName = csvFile.Split(delimiter)[csvFile.Split(delimiter).Length - 1];

                do
                {
                    try
                    {
                        firstLine = File.ReadLines(csvFile).First();
                    }
                    catch (IOException)
                    {
                        Console.Util.WriteMessageInCenter(fileName + " is currently open.  Press any key once the file has been closed.", ConsoleColor.Red);
                        System.Console.ReadKey();
                        System.Console.Write("\b \b ");
                    }
                } while (firstLine == string.Empty) ;

                if (MatchesMicroContents(firstLine))
                    MicroFilePath = csvFile;

                if (MatchesTitrationContents(firstLine))
                    TitrationFilePath = csvFile;
            }
        }
        /// <summary>
        /// Checks if the string contains expected values from micro data
        /// </summary>
        /// <param name="firstLineOfFile"></param>
        /// <returns></returns>
        private bool MatchesMicroContents(string firstLineOfFile)
        {
            string[] delimitedLine = firstLineOfFile.Split(new char[] { ',' });

            if (delimitedLine[0] == "H1" || delimitedLine[0] == "L1" || delimitedLine[0] == "S1" || delimitedLine[0] == "D1")
                return true;
            else
                return false;
        }
        /// <summary>
        /// Checks if the string contains expected values from titration data
        /// </summary>
        /// <param name="firstLineOfFile"></param>
        /// <returns></returns>
        private bool MatchesTitrationContents(string firstLineOfFile)
        {
            string[] delimitedLine = firstLineOfFile.Split(new char[] { ',' });

            if (delimitedLine[2] == "H1" || delimitedLine[2] == "L1" || delimitedLine[2] == "S1" || delimitedLine[2] == "D1")
                return true;
            else
                return false;
        }
        public List<int> FindTitrationIndices(string recipeCode, DateTime madeDate, string factoryCode)
        {
            List<int> indices = new List<int>();
            string jobNumber = string.Empty;
            string madeDateAsString = madeDate.ToShortDateString();
            string madeDateAsTwoDigitYearString = madeDate.ToString("M/d/yy");

            for (int i = 0; i < DelimitedTitrationResults.Count; i++)
            {
                if (DelimitedTitrationResults[i][2] == factoryCode && (DelimitedTitrationResults[i][0] == madeDateAsString ||
                    (DelimitedTitrationResults[i][0] == madeDateAsTwoDigitYearString)) && DelimitedTitrationResults[i][4] == recipeCode)
                {
                    indices.Add(i);

                    if (string.IsNullOrEmpty(jobNumber))
                    {
                        jobNumber = DelimitedTitrationResults[i][3];
                    }
                }
            }

            // Runs a second check under a different factory
            if (indices.Count == 0)
            {
                if (factoryCode == "H1")
                    factoryCode = "L1";
                else if (factoryCode == "L1")
                    factoryCode = "H1";

                for (int i = 0; i < DelimitedTitrationResults.Count; i++)
                {
                    if (DelimitedTitrationResults[i][2] == factoryCode && (DelimitedTitrationResults[i][0] == madeDateAsString ||
                        (DelimitedTitrationResults[i][0] == madeDateAsTwoDigitYearString)) && DelimitedTitrationResults[i][4] == recipeCode)
                    {
                        indices.Add(i);

                        if (string.IsNullOrEmpty(jobNumber))
                        {
                            jobNumber = DelimitedTitrationResults[i][3];
                        }
                    }
                }
            }
            return indices;
        }
        public float GetTitrationValue(List<int> indices, TitrationOffset offset)
        {
            List<float> results = new List<float>();

            foreach (int index in indices)
            {
                for (int i = 0; i < DelimitedTitrationResults[index].Count; i++)
                {
                    string value = DelimitedTitrationResults[index][i];

                    if (value == "Original" || value == "ReTest_1" || value == "ReTest_2" || value == "ReTest_3" || value == "ReTest_4" || value == "ReTest_5")
                    {
                        if (DelimitedTitrationResults[index][i + (int)offset] != "*")
                            results.Add(Convert.ToSingle(DelimitedTitrationResults[index][i + (int)offset]));
                    }
                }
            }

            float sum = 0;

            foreach (float value in results)
            {
                sum += value;
            }

            float average = sum / results.Count;

            int closestValueIndex = 0;
            float smallestDifference = 10000000000;
            float currentDifference;

            for (int i = 0; i < results.Count; i++)
            {
                currentDifference = results[i] - average >= 0 ? results[i] - average : average - results[i];

                if (currentDifference < smallestDifference)
                {
                    smallestDifference = currentDifference;
                    closestValueIndex = i;
                }
            }
            return results.Count == 0 ? -1 : results[closestValueIndex];
        }
    }
}
