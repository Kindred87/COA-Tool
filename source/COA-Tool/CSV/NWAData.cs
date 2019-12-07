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
using CoA_Tool.Utility;

namespace CoA_Tool.CSV
{
    /// <summary>
    /// Represents CSV dumps from NWA
    /// </summary>
    class NWAData
    {
        // Class variables
        //  Lists
        public List<List<string>> DelimitedMicroResults = new List<List<string>>();
        public List<List<string>> DelimitedTitrationResults = new List<List<string>>();
        /// <summary>
        /// Represents the number of columns to the right of the original/re-test value a titration value can be found
        /// </summary>
        public enum TitrationOffset { Acidity = 5, ViscosityCPS = 2, ViscosityCM = 3, Salt = 4, pH = 6}
        public enum MicroOffset { Yeast = 9, Mold = 11, Aerobic = 15, Coliform = 7, Lactic = 13, EColi = 5 }

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
                    ConsoleOps.WriteMessageInCenter("Could not locate micro data on the desktop." +  
                        "  Press a key once the file is ready to be loaded.", ConsoleColor.Red);
                    Console.ReadKey();
                    Console.Write("\b \b ");
                }
                if(titrationPath == string.Empty)
                {
                    ConsoleOps.WriteMessageInCenter("Could not locate titration data on the desktop." + 
                        "  Press a key once the file is ready to be loaded.", ConsoleColor.Red);
                    Console.ReadKey();
                    Console.Write("\b \b ");
                }
            }

            if (FileOutdated(microPath, out int hoursSinceMicroUpdated))
            {
                if (new ConsoleInteraction.SelectionMenu(new string[] { "Reload", "Continue regardless" }.ToList(), 
                    "Reload file?", 
                    "Micro data hasn't been updated in " + hoursSinceMicroUpdated + " hours.").UserChoice == "Reload")
                {
                    new ConsoleInteraction.SelectionMenu(new string[] { "File has been updated" }.ToList(),
                        "File ready?",
                        "Update the micro data before continuing.");
                    LoadCSVFiles();
                }
            }
            
            if(FileOutdated(titrationPath, out int hoursSinceTitrationUpdated))
            {
                if(new ConsoleInteraction.SelectionMenu(new string[] { "Reload", "Continue regardless"}.ToList(),
                    "Reload file?",
                    "Titration data hasn't been updated in " + hoursSinceTitrationUpdated + " hours.").UserChoice == "Reload")
                {
                    new ConsoleInteraction.SelectionMenu(new string[] { "File has been updated" }.ToList(),
                        "File ready?",
                        "Update the Titration data before continuing.");
                    LoadCSVFiles();
                }
            }

            ConsoleOps.WriteMessageInCenter("Loading micro data...");
            foreach (string line in File.ReadLines(microPath))
            {
                DelimitedMicroResults.Add(line.Split(new char[] { ',' }).ToList());
            }
            ConsoleOps.WriteMessageInCenter("Loading titration data...");
            foreach (string line in File.ReadLines(titrationPath))
            {
                DelimitedTitrationResults.Add(line.Split(new char[] { ',' }).ToList());
            }

            ConsoleOps.RemoveMessageInCenter();
        }

        // Private methods
        /// <summary>
        /// Prompts user to optionally reload files if last update meets or exceeds 6 hours
        /// </summary>
        /// <param name="microPath"></param>
        /// <param name="titrationPath"></param>
        /// <returns></returns>
        private bool FileOutdated(string path, out int hoursSinceUpdate)
        {
            hoursSinceUpdate = (int)(DateTime.Now - File.GetLastWriteTime(path)).TotalHours;

            if (hoursSinceUpdate >= 6)
            {
                
                return true;
            }
            else
            {
                return false;
            }
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
                        ConsoleOps.WriteMessageInCenter(fileName + " is currently open.  Press any key once the file has been closed.", ConsoleColor.Red);
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
        /// <summary>
        /// Retrieves indices of entries in DelimitedTitrationResults matching the given criteria 
        /// </summary>
        /// <param name="recipeCode"></param>
        /// <param name="madeDate"></param>
        /// <param name="factoryCode"></param>
        /// <returns></returns>
        public List<int> TitrationIndices(string recipeCode, DateTime madeDate, string factoryCode)
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
        /// <summary>
        /// Retrieves the appropriate values from the delimited titration results.
        /// </summary>
        /// <param name="searchIndices">The indices in the delimited micro results to parse.</param>
        /// <param name="offset">Signifies the column to search within.</param>
        /// <returns></returns>
        public List<string> TitrationValues(List<int> searchIndices, TitrationOffset offset)
        {
            List<string> rawTitrationValues = new List<string>();

            foreach (int index in searchIndices)
            {
                for (int i = 0; i < DelimitedTitrationResults[index].Count; i++)
                {
                    string value = DelimitedTitrationResults[index][i];

                    if (value == "Original" || value == "ReTest_1" || value == "Re_Test_1" || value == "Re_Test_2" || value == "Re_Test_3")
                    {
                        if (DelimitedTitrationResults[index][i + (int)offset] != "*")
                        {
                            rawTitrationValues.Add(DelimitedTitrationResults[index][i + (int)offset]);
                        }
                    }
                }
            }

            return rawTitrationValues;
        }
        /// <summary>
        /// Retrieves indices of entries in DelimitedTitrationResults matching the given criteria 
        /// </summary>
        /// <param name="recipeCode">The five-digit recipe code to search under</param>
        /// <param name="lotCode">The twelve-digit lot code to search under</param>
        /// <param name="madeDate">The product's manufacturing date</param>
        /// <returns></returns>
        public List<int> MicroIndices(string recipeCode, string lotCode, DateTime madeDate)
        {
            List<int> indices = new List<int>();

            string productCode = Lot.ProductCode(lotCode);
            string factoryCode = Lot.FactoryCode(lotCode);
            string madeDateAsString = madeDate.ToShortDateString();
            string madeDateAsTwoDigitYearString = madeDate.ToString("M/d/yy");

            for (int i = 0; i < DelimitedMicroResults.Count; i++)
            {
                if (DelimitedMicroResults[i][0] == factoryCode && DelimitedMicroResults[i][7] == recipeCode && DelimitedMicroResults[i][10] == productCode &&
                    (madeDateAsString == DelimitedMicroResults[i][9] || madeDateAsTwoDigitYearString == DelimitedMicroResults[i][9]))
                {
                    indices.Add(i);
                }
            }
            return indices;
        }
        /// <summary>
        /// Retrieves indices of entries in DelimitedMicroResults matching the given criteria 
        /// </summary>
        /// <param name="productCode">The five-digit product code to search under</param>
        /// <param name="madeDate">The product's manufacturing date</param>
        /// <param name="supplier">Value to look for in the supplier column</param>
        /// <returns></returns>
        public List<int> MicroIndices(string productCode, DateTime madeDate, string supplier)
        {
            List<int> indices = new List<int>();

            string madeDateAsString = madeDate.ToShortDateString();
            string madeDateAsTwoDigitYearString = madeDate.ToString("M/d/yy");

            for (int i = 0; i < DelimitedMicroResults.Count; i++)
            {
                if (DelimitedMicroResults[i][16] == supplier && DelimitedMicroResults[i][10] == productCode &&
                    (madeDateAsString == DelimitedMicroResults[i][9] || madeDateAsTwoDigitYearString == DelimitedMicroResults[i][9]))
                {
                    indices.Add(i);
                }
            }
            return indices;
        }
        /// <summary>
        /// Retrieves the appropriate values from the delimited micro results.
        /// </summary>
        /// <param name="searchIndices">The indices in the delimited micro results to parse.</param>
        /// <param name="offset">Signifies the column to search within.</param>
        /// <returns></returns>
        public List<string> MicroValues(List<int> searchIndices, MicroOffset offset)
        {
            List<string> rawMicroValues = new List<string>();

            foreach (int searchIndex in searchIndices)
            {
                for (int columnIterator = 0; columnIterator < DelimitedMicroResults[searchIndex].Count; columnIterator++)
                {
                    if (DelimitedMicroResults[searchIndex][columnIterator] == "HURRICANE" || DelimitedMicroResults[searchIndex][columnIterator] == "Hurricane" || DelimitedMicroResults[searchIndex][columnIterator] == "Lowell" ||
                        DelimitedMicroResults[searchIndex][columnIterator] == "Sandpoint")
                    {
                        string rawValue = DelimitedMicroResults[searchIndex][columnIterator + (int)offset];
                        
                        if (string.IsNullOrEmpty(rawValue))
                        {
                            continue;
                        }
                        else
                        {
                            rawMicroValues.Add(rawValue);
                        }
                    }
                }
            }

            return rawMicroValues;
        }
        /// <summary>
        /// Filters out-of-spec values provided that at least one value is in-spec, otherwise the provided list is returned
        /// </summary>
        /// <param name="unsortedValues"></param>
        /// <param name="offset"></param>
        /// <returns></returns>
        public string ProductCodeFromMicroIndex(int index)
        {
            return DelimitedMicroResults[index][10];
        }
        public string RecipeCodeFromMicroIndex(int index)
        {
            return DelimitedMicroResults[index][7];
        }
        public string RecipeCodeFromTitrationIndex(int index)
        {
            return DelimitedTitrationResults[index][4];
        }
        /// <summary>
        /// Retrieves batch values from a collection of indices in micro results in a comma-containing string
        /// </summary>
        /// <param name="indices">The target indices in DelimitedTitrationResults</param>
        /// <returns></returns>
        public string BatchValuesFromTitrationIndices(List<int> indices)
        {
            List<string> batchValues = new List<string>();
            string indexValue = "";
            int offset = 0;
            bool continueLoop = true;

            foreach(int index in indices)
            {
                while (continueLoop)
                {
                    indexValue = DelimitedTitrationResults[index][13 + offset];

                    if(indexValue == "ReTest_1" || indexValue == "Original" || indexValue == "Re_Test_1" || indexValue == "Re_Test_2" || indexValue == "Re_Test_3")
                    {
                        continueLoop = false;
                    }
                    else
                    {
                        batchValues.Add(indexValue);
                        offset++;
                    }
                }
            }

            string batchesCombinedInString = "";

            for(int i = 0; i < batchValues.Count; i++)
            {
                if(batchValues[i] != "*")
                {
                    batchesCombinedInString += batchValues[i];

                    if (i + 1 < batchValues.Count && batchValues[i + 1] != "*")
                    {
                        batchesCombinedInString += ", ";
                    }
                }
            }

            return batchesCombinedInString;
        }
        /// <summary>
        /// Retrieves batch values from a collection of indices in micro results in a comma-containing string
        /// </summary>
        /// <param name="indices">The target indices in DelimitedMicroResults</param>
        /// <returns></returns>
        public string BatchValuesFromMicroIndices(List<int> indices)
        {
            List<string> batchValues = new List<string>();
            string indexValue = "";
            int offset = 0;
            bool continueLoop = true;

            foreach (int index in indices)
            {
                while (continueLoop)
                {
                    indexValue = DelimitedMicroResults[index][11 + offset];

                    if (DateTime.TryParse(indexValue, out DateTime dateTime) || indexValue == "*")
                    {
                        continueLoop = false;
                    }
                    else
                    {
                        batchValues.Add(indexValue);
                        offset++;
                    }
                }
            }

            string batchesCombinedInString = "";

            for (int i = 0; i < batchValues.Count; i++)
            {
                batchesCombinedInString += batchValues[i];

                if (i + 1 < batchValues.Count)
                {
                    batchesCombinedInString += ", ";
                }
            }

            return batchesCombinedInString;
        }
        public List<string> WaterActivityValues(List<int> searchIndices)
        {
            List<string> rawWaterActivityValues = new List<string>();

            foreach (int rowIndex in searchIndices)
            {
                for (int columnIterator = 0; columnIterator < DelimitedTitrationResults[rowIndex].Count; columnIterator++)
                {
                    string coordinateValue = DelimitedTitrationResults[rowIndex][columnIterator];

                    if (coordinateValue == "Original" || coordinateValue == "ReTest_1" || coordinateValue == "Re_Test_1" || coordinateValue == "Re_Test_2" || coordinateValue == "Re_Test_3")
                    {
                        string rawValue = "";

                        // Target value is a float formatted as 0.### or .###.  Value can potentially be 1.0 (100% water activity).  0.0 is, in all practicality, impossible.
                        // Default value is "*", which is synonymous with N/A
                        if (DelimitedTitrationResults[rowIndex][columnIterator + 12].Contains('.'))
                        {
                            rawValue += DelimitedTitrationResults[rowIndex][columnIterator + 12];
                        }

                        // Two different sub-indices are targeted due to possible inclusion of commas in the string splitting it into two separate
                        // columns within the CSV file
                        if (DelimitedTitrationResults[rowIndex][columnIterator + 13].Contains('.'))
                        {
                            rawValue += DelimitedTitrationResults[rowIndex][columnIterator + 13];
                        }

                        if(rawValue.Length > 0)
                        {
                            rawWaterActivityValues.Add(rawValue);
                        }
                    }
                }
            }
            return rawWaterActivityValues;
        }
    }
}
