using System;
using System.Collections.Generic;
using System.Text;
/// <summary>
/// Contains static collections of methods
/// </summary>
namespace CoA_Tool.Utility
{
    /// <summary>
    /// Static collection of methods pertaining to lot codes
    /// </summary>
    static class Lot
    {
        /// <summary>
        /// Retrieves product code from a given lot
        /// </summary>
        /// <param name="lot">The lot containing the product code</param>
        /// <returns></returns>
        public static string ProductCode(string lotCode)
        {
            string productCode = lotCode[0].ToString();
            productCode += lotCode[1];
            productCode += lotCode[2];
            productCode += lotCode[3];
            productCode += lotCode[4];

            return productCode;
        }
        /// <summary>
        /// Determines manufacturing site for a given lot
        /// </summary>
        /// <param name="lotCode">The target lot</param>
        /// <returns></returns>
        public static string ManufacturingSite(string lotCode)
        {
            string locationCode = lotCode[5].ToString();
            locationCode += lotCode[6];

            if (locationCode == "01")
            {
                return "Sandpoint, ID";
            }
            else if (locationCode == "02")
            {
                return "Lowell, MI";
            }
            else if (locationCode == "03")
            {
                return "Hurricane, UT";
            }
            else
            {
                return locationCode;
            }
        }
        /// <summary>
        /// Determines the alphanumeric factory code from the provided lot
        /// </summary>
        /// <param name="lotCode">The lot from which to determine the factory code</param>
        /// <returns></returns>
        public static string FactoryCode(string lotCode)
        {
            string locationCode = lotCode[5].ToString();
            locationCode += lotCode[6];

            if (locationCode == "01")
                return "S1";
            else if (locationCode == "02")
                return "L1";
            else if (locationCode == "03")
                return "H1";
            else
                return string.Empty;
        }
        /// <summary>
        /// Converts best-by-related characters in a given lot code to a DateTime object.  The return value
        ///  indicates whether the conversion succeeded.
        /// </summary>
        /// <param name="lotCode">The 12-digit lot code containing the best by to be retrieved</param>
        /// <param name="bestBy">DateTime representing the expiry date of product</param>
        /// <returns></returns>
        public static bool TryParseBestBy(string lotCode, out DateTime bestByDate)
        {
            // 00000 00 000111  (Two digit values arranged vertically)
            // 01234 56 789012

            // 16254 02 021520 (Lot code representation)

            bool conversionSucceeded = true; // Assigned to false if any Int32.TryParse returns false

            int parsedValue; // Used to modify values *before* variable assignment
            
            int bestByMonth = 0; 
            if(Int32.TryParse(lotCode[7].ToString(), out parsedValue) == true) // First digit of bestByMonth
            {
                bestByMonth = parsedValue * 10;
            }
            else
            {
                conversionSucceeded = false;
            }

            if(Int32.TryParse(lotCode[8].ToString(), out parsedValue) == true) // Second digit of bestByMonth
            {
                bestByMonth += parsedValue;
            }
            else
            {
                conversionSucceeded = false;
            }


            int bestByDay = 0;
            if(Int32.TryParse(lotCode[9].ToString(), out parsedValue) == true) // First digit of bestByDay
            {
                bestByDay = parsedValue * 10;
            }
            else
            {
                conversionSucceeded = false;
            }

            if(Int32.TryParse(lotCode[10].ToString(), out parsedValue) == true) // Second digit of bestByDay
            {
                bestByDay += parsedValue;
            }

            int bestByYear = 0;
            if(Int32.TryParse(lotCode[11].ToString(), out parsedValue) == true) // Third digit of bestByYear
            {
                bestByYear = 2000 + parsedValue * 10;
            }
            else
            {
                conversionSucceeded = false;
            }

            if(Int32.TryParse(lotCode[12].ToString(), out parsedValue) == true) // Fourth digit of bestByYear
            {
                bestByYear += parsedValue;
            }
            else
            {
                conversionSucceeded = false;
            }

            bestByDate = new DateTime(bestByYear, bestByMonth, bestByDay);

            return conversionSucceeded;
        }
    }
}
