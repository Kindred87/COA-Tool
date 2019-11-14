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
    }
}
