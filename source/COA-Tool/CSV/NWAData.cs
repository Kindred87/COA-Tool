﻿using System;
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
        /// Searches for titration values at the provided indices.  The return value
        ///  indicates whether a valid value was located.
        /// </summary>
        /// <param name="indices"></param>
        /// <param name="offset"></param>
        /// <param name="titrationValue"></param>
        /// <returns></returns>
        public bool TitrationValueExists(List<int> indices, TitrationOffset offset, out float titrationValue) 
        {
            List<float> results = new List<float>();

            foreach (int index in indices)
            {
                for (int i = 0; i < DelimitedTitrationResults[index].Count; i++)
                {
                    string value = DelimitedTitrationResults[index][i];

                    if (value == "Original" || value == "ReTest_1" || value == "Re_Test_1" || value == "Re_Test_2" || value == "Re_Test_3")
                    {
                        if (DelimitedTitrationResults[index][i + (int)offset] != "*")
                        {
                            results.Add(Convert.ToSingle(DelimitedTitrationResults[index][i + (int)offset]));
                        }
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
            
            if(results.Count > 0)
            {
                titrationValue = results[closestValueIndex];
                return true;
            }
            else
            {
                titrationValue = 0;
                return false;
            }
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
        public int GetMicroValue(List<int> indices, MicroOffset offset)
        {
            List<int> microValues = new List<int>();

            foreach (int index in indices)
            {
                for (int i = 0; i < DelimitedMicroResults[index].Count; i++)
                {
                    if (DelimitedMicroResults[index][i] == "HURRICANE" || DelimitedMicroResults[index][i] == "Hurricane" || DelimitedMicroResults[index][i] == "Lowell" || DelimitedMicroResults[index][i] == "Sandpoint")
                    {
                        if (string.IsNullOrEmpty(DelimitedMicroResults[index][i + (int)offset]) || DelimitedMicroResults[index][i + (int)offset] == "*")
                            continue;
                        else

                            microValues.Add(Convert.ToInt32(DelimitedMicroResults[index][i + (int)offset].Trim()));
                    }
                }
            }

            int largestValue = -1;

            foreach (int value in FilterMicroValues(microValues, offset))
            {
                if (value > largestValue)
                {
                    largestValue = value;
                }
            }

            return largestValue;
        }
        /// <summary>
        /// Filters out-of-spec values provided that at least one value is in-spec, otherwise the provided list is returned
        /// </summary>
        /// <param name="unsortedValues"></param>
        /// <param name="offset"></param>
        /// <returns></returns>
        private List<int> FilterMicroValues(List<int> unsortedValues, MicroOffset offset)
        {
            List<int> sortedValues = new List<int>();

            foreach (int value in unsortedValues)
            {
                if (offset == MicroOffset.Aerobic && value < 100000) // 100k
                    sortedValues.Add(value);
                else if (offset == MicroOffset.Coliform && value < 100)
                    sortedValues.Add(value);
                else if (offset == MicroOffset.Lactic && value < 1000)
                    sortedValues.Add(value);
                else if (offset == MicroOffset.Mold && value < 1000)
                    sortedValues.Add(value);
                else if (offset == MicroOffset.Yeast && value < 1000)
                    sortedValues.Add(value);
                else if (offset == MicroOffset.EColiform && value == 0)
                    sortedValues.Add(value);
            }

            if (sortedValues.Count != 0)
            {
                return sortedValues;
            }
            else
            {
                return unsortedValues;
            }

        }
        public bool MicroValueInSpec(int value, MicroOffset offset)
        {
            if (offset == MicroOffset.Aerobic)
            {
                if (value < 100000)
                    return true;
                else
                    return false;
            }
            else if (offset == MicroOffset.Coliform)
            {
                if (value < 100)
                    return true;
                else
                    return false;
            }
            else if (offset == MicroOffset.Lactic)
            {
                if (value < 1000)
                    return true;
                else
                    return false;
            }
            else if (offset == MicroOffset.Mold)
            {
                if (value < 1000)
                    return true;
                else
                    return false;
            }
            else // when (offset == MicroOffset.Yeast)
            {
                if (value < 1000)
                    return true;
                else
                    return false;
            }
        }
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
                batchesCombinedInString += batchValues[i];

                if(i + 1 < batchValues.Count)
                {
                    batchesCombinedInString += ", ";
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
        /// <summary>
        /// Searches for water activity values at the provided indices.  The return value
        ///  indicates whether a value was located.  Secondary out boolean indicates whether the value was valid.
        /// </summary>
        /// <param name="indices"></param>
        /// <param name="waterActivity"></param>
        /// <param name="valueAlsoValid"></param>
        /// <returns></returns>
        public bool WaterActivityExists(List<int> indices, out float waterActivity, out bool valueAlsoValid)
        {
            List<float> validWaterActivityValues = new List<float>();

            foreach (int index in indices)
            {
                for (int i = 0; i < DelimitedTitrationResults[index].Count; i++)
                {
                    string value = DelimitedTitrationResults[index][i];

                    if (value == "Original" || value == "ReTest_1" || value == "Re_Test_1" || value == "Re_Test_2" || value == "Re_Test_3")
                    {
                        string lineToParse = "";

                        // Target value is a float formatted as 0.### or .###.  Value can potentially be 1.0 (100% water activity).  0.0 is, in all practicality, impossible.
                        // Default value is "*", which is synonymous with N/A
                        if (DelimitedTitrationResults[index][i + 12].Contains('.')) 
                        {
                            lineToParse += DelimitedTitrationResults[index][i + 12];
                        }
                        
                        // Two different sub-indices are targeted due to possible inclusion of commas in the string splitting it into two separate
                        // columns within the CSV file
                        if (DelimitedTitrationResults[index][i + 13].Contains('.')) 
                        {
                            lineToParse += DelimitedTitrationResults[index][i + 13];
                        }

                        if (lineToParse.Length > 0) // True if a decimal was found.  
                        {
                            float validValueForWaterActivity = 0; // Water activity is typically a positive value less than 1

                            // delimitedLine can potentially contain more than two indices, though it's unlikely.
                            // Following code intended to take the above scenario into account.
                            List<string> delimitedLine = lineToParse.Split(new char[] { '.' }).ToList();

                            for (int lineIndex = 0; lineIndex < delimitedLine.Count; lineIndex++)
                            {
                                bool addToList = false;

                                if (delimitedLine[lineIndex].Count() < 3)
                                {
                                    continue;
                                }

                                // Target float has three digits following the decimal
                                if (Char.IsDigit(delimitedLine[lineIndex][0]) && Char.IsDigit(delimitedLine[lineIndex][1]) && Char.IsDigit(delimitedLine[lineIndex][2]))
                                {
                                    addToList = true;

                                    if (lineIndex != 0 && delimitedLine[lineIndex - 1].Last() == '1')
                                    {
                                        validValueForWaterActivity = 1;
                                    }

                                    for (int digitCount = 1; digitCount <= 3; digitCount++) // Assigns non-integer value, if possible.  
                                    {
                                        if (Single.TryParse(delimitedLine[1][digitCount - 1].ToString(), out float parsedValue) == true)
                                        {
                                            validValueForWaterActivity += parsedValue / (float)Math.Pow(10, digitCount); // Add-assigns each non-integer value by dividing by increasing multiples of 10
                                        }
                                    }
                                }
                                if (addToList)
                                {
                                    validWaterActivityValues.Add(validValueForWaterActivity);
                                }
                            }
                            
                        }
                    }
                }
            }
            
            float sum = 0;

            foreach (float value in validWaterActivityValues)
            {
                sum += value;
            }

            float average = sum / validWaterActivityValues.Count;

            int nearestToAverageIndex = 0;
            float smallestDifference = 10000000000;
            float currentDifference;

            for (int i = 0; i < validWaterActivityValues.Count; i++)
            {
                currentDifference = validWaterActivityValues[i] - average >= 0 ? validWaterActivityValues[i] - average : average - validWaterActivityValues[i];

                if (currentDifference < smallestDifference)
                {
                    smallestDifference = currentDifference;
                    nearestToAverageIndex = i;
                }
            }

            if (validWaterActivityValues.Count > 0)
            {
                waterActivity = validWaterActivityValues[nearestToAverageIndex];
                valueAlsoValid = true;
                return true;
            }
            else
            {
                waterActivity = 0;
                valueAlsoValid = false;
                return false;
            }
        }
    }
}
