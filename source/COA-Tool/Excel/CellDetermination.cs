using System;
using System.Collections.Generic;
using System.Drawing;
using System.Linq;
using System.Text;

namespace CoA_Tool.Excel
{
    /// <summary>
    /// Contains methods used to determine appropriate cell values and attributes.
    /// </summary>
    static class CellDetermination
    {
        /// <summary>
        /// Determines the string value of a cell for micro results in the worksheet.
        /// </summary>
        /// <param name="rawMicroValues"></param>
        /// <param name="microType"></param>
        /// <returns></returns>
        public static string ValueForMicroCell(List<string> rawMicroValues, CSV.NWAData.MicroOffset microType, string cityOfManufacture)
        {
            if (rawMicroValues.Count == 0)
            {
                return "Search error";
            }

            List<int> numericalMicroValues = OnlyIntegersFrom(rawMicroValues);

            if (numericalMicroValues.Count > 0)
            {
                int greatestValue = 0;

                foreach (int value in numericalMicroValues)
                {
                    if (value > greatestValue)
                    {
                        greatestValue = value;
                    }
                }

                if (greatestValue == 0)
                {
                    return XMLOperations.Definitions.NodeValueViaXPath(XMLOperations.Definitions.DefinitionFiles.Micro,
                        "/Dilutions_By_Factory/factory[@name = '" + cityOfManufacture + "']/" + Convert.ToString(microType));
                }
                else

                {
                    return greatestValue.ToString();
                }
            }
            else
            {
                return "N/A";
            }
        }
        /// <summary>
        /// Determines the string value of a cell for titration results in the worksheet.
        /// </summary>
        /// <param name="rawTitrationValues"></param>
        /// <param name="titrationTestCategory"></param>
        /// <returns></returns>
        public static string ValueForTitrationCell(List<string> rawTitrationValues, CSV.NWAData.TitrationOffset titrationTestCategory)
        {
            if(rawTitrationValues.Count == 0)
            {
                return "Search error";
            }

            List<float> numericalTitrationValues = OnlyFloatsFrom(rawTitrationValues);

            if(numericalTitrationValues.Count > 0)
            {
                return ValueNearestToMean(numericalTitrationValues).ToString();
            }
            else
            {
                return "N/A";
            }
        }
        /// <summary>
        /// Determines the string value of a cell for water activity in the worksheet.
        /// </summary>
        /// <param name="rawWaterActivityValues"></param>
        /// <returns></returns>
        public static string ValueForWaterActivityCell(List<string> rawWaterActivityValues)
        {
            if(rawWaterActivityValues.Count == 0)
            {
                return "Search error";
            }

            List<float> waterActivityValues = new List<float>();

            foreach(string rawValue in rawWaterActivityValues)
            {
                if(TryExtractWaterActivityFloat(rawValue, out float extractedValue))
                {
                    waterActivityValues.Add(extractedValue);
                }
            }

            if(waterActivityValues.Count > 0)
            {
                return ValueNearestToMean(waterActivityValues).ToString();
            }
            else
            {
                return "N/A";
            }
        }
        /// <summary>
        /// Indicates if a string contains a float value with three fractional digits.
        /// </summary>
        /// <param name="stringToParse">The string in which to seek the desired value.</param>
        /// <param name="extractedValue">The water activity value extracted from the string.  Is 0 if parsing failed.</param>
        /// <returns></returns>
        private static bool TryExtractWaterActivityFloat(string stringToParse, out float extractedValue)
        {
            List<string> delimitedLine = stringToParse.Split(new char[] { '.' }).ToList();
            bool valueExtracted = false;
            extractedValue = 0;

            for (int lineIndex = 0; lineIndex < delimitedLine.Count; lineIndex++)
            {
                if (delimitedLine[lineIndex].Count() < 3)
                {
                    continue;
                }

                // Target float has three fractional digits
                if (Char.IsDigit(delimitedLine[lineIndex][0]) && Char.IsDigit(delimitedLine[lineIndex][1]) && Char.IsDigit(delimitedLine[lineIndex][2]))
                {
                    valueExtracted = true;

                    // If there are decimal values but the integer value is one
                    if (lineIndex != 0 && delimitedLine[lineIndex - 1].Last() == '1')
                    {
                        extractedValue = 1; // Since a value of 1 is the maximum, the value is set to 1.0
                    }

                    for (int digitCount = 1; digitCount <= 3; digitCount++) // Assigns non-integer value, if possible.  
                    {
                        if (Single.TryParse(delimitedLine[1][digitCount - 1].ToString(), out float parsedValue) == true)
                        {
                            extractedValue += parsedValue / (float)Math.Pow(10, digitCount); // Add-assigns each non-integer value by dividing by increasing multiples of 10
                        }
                        else
                        {
                            valueExtracted = false;
                        }
                    }
                }
                else
                {
                    valueExtracted = false;
                }
            }
            return valueExtracted;
        }
        /// <summary>
        /// Returns the converted values of integer-convertable strings.
        /// </summary>
        /// <param name="targetValues">The strings to attempt integer conversion with.</param>
        /// <returns></returns>
        private static List<int> OnlyIntegersFrom(List<string> targetValues)
        {
            List<int> integerParseableValues = new List<int>();

            foreach (string item in targetValues)
            {
                if (Int32.TryParse(item, out int result)) // Result is left unused
                {
                    integerParseableValues.Add(result);
                }
            }

            return integerParseableValues;
        }
        /// <summary>
        /// Returns the converted values of integer-convertable strings.
        /// </summary>
        /// <param name="targetValues">The strings to attempt integer conversion with.</param>
        /// <returns></returns>
        private static List<float> OnlyFloatsFrom(List<string> targetValues)
        {
            List<float> floatParseableValues = new List<float>();

            foreach(string item in targetValues)
            {
                if(Single.TryParse(item, out float result))
                {
                    floatParseableValues.Add(result);
                }
            }
            return floatParseableValues;
        }
        /// <summary>
        /// Returns a Color dependent on the target cell's value and comment.
        /// </summary>
        /// <param name="targetWorksheet">The worksheet object containing the target cell.</param>
        /// <param name="targetCellRow">Represents the row in which the target cell is positioned.</param>
        /// <param name="targetCellColumn">Represents the column in which the target cell is positioned.</param>
        /// <returns></returns>
        public static Color FontColorForCell(string cellValue, string commentText)
        {
            if (cellValue == "Search error")
            {
                return Color.Red;
            }
            else if (commentText.Contains("invalid"))
            {
                return Color.OrangeRed;
            }
            else
            {
                return Color.Black;
            }
        }
        /// <summary>
        /// Returns a Color dependent on the target cell's value.
        /// </summary>
        /// <param name="targetWorksheet">The worksheet object containing the target cell.</param>
        /// <param name="targetCellRow">Represents the row in which the target cell is positioned.</param>
        /// <param name="targetCellColumn">Represents the column in which the target cell is positioned.</param>
        /// <returns></returns>
        public static Color FontColorForCell(string cellValue)
        {
            if (cellValue == "Search error")
            {
                return Color.Red;
            }
            else
            {
                return Color.Black;
            }
        }
        /// <summary>
        /// Determines if a comment is required for a particular micro or titration cell.
        /// </summary>
        /// <param name="rawValues">The values relevant to the cell.</param>
        /// <param name="commentText">The value of the comment to be written.</param>
        /// <returns></returns>
        public static bool CommentNeededForMicroOrTitration(List<string> rawValues, out string commentText)
        {
            foreach (string rawValue in rawValues)
            {
                if (rawValue != "*" && (Int32.TryParse(rawValue, out _) == false && Single.TryParse(rawValue, out _) == false))
                {
                    commentText = "One or more values were invalid";
                    return true;
                }
            }
            commentText = "";
            return false;
        }
        /// <summary>
        /// Determines if a comment is required for a particular water activity cell.
        /// </summary>
        /// <param name="rawValues">The values relevant to the cell.</param>
        /// <param name="commentText">The value of the comment to be written.</param>
        /// <returns></returns>
        public static bool CommentNeededForWaterActivity(List<string> rawValues, out string commentText)
        {
            foreach(string rawValue in rawValues)
            {
                if(TryExtractWaterActivityFloat(rawValue, out _) == false)
                {
                    commentText = "One or more values were invalid";
                    return true;
                }
            }
            commentText = "";
            return false;
        }
        /// <summary>
        /// Determines which of the provided values is closest to the mean
        /// </summary>
        /// <param name="providedValues"></param>
        /// <returns></returns>
        static private float ValueNearestToMean(List<float> providedValues)
        {
            float sum = 0;

            foreach (float value in providedValues)
            {
                sum += value;
            }

            float average = sum / providedValues.Count;

            int nearestToAverageIndex = 0;
            float smallestDifference = 10000000000;
            float currentDifference;

            for (int i = 0; i < providedValues.Count; i++)
            {
                currentDifference = providedValues[i] - average >= 0 ? providedValues[i] - average : average - providedValues[i];

                if (currentDifference < smallestDifference)
                {
                    smallestDifference = currentDifference;
                    nearestToAverageIndex = i;
                }
            }

            return providedValues[nearestToAverageIndex];
        }
    }
}
