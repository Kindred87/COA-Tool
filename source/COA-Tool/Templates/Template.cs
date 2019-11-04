using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.IO;
using CoA_Tool;


namespace CoA_Tool
{
    /// <summary>
    /// Contains information on how the Excel document is populated
    /// </summary>
    class Template
    {
        // Enums
        public enum Algorithm { Standard, ResultsFromDateOnwards}
        private enum ContentCategories { None, Algorithm, MainContentBlock, }

        // Objects
        public Console.SelectionMenu Menu { get; private set; }

        // Enum variables
        public Algorithm SelectedAlgorithm;

        // Bools
        public bool IncludeCustomerName;
        public bool IncludeSalesOrder;
        public bool IncludePurchaseOrder;
        public bool IncludeGenerationDate;
        public bool IncludeProductName;
        public bool IncludeRecipeAndItem;
        public bool IncludeLotCode;
        public bool IncludeBatch;
        public bool IncludeBestByDate;
        public bool IncludeManufacturingSite;
        public bool IncludeManufacturingDate;
        public bool IncludeAcidity;
        public bool IncludepH;
        public bool IncludeViscosityCM;
        public bool IncludeViscosityCPS;
        public bool IncludeWaterActivity;
        public bool IncludeBrixSlurry;
        public bool IncludeYeast;
        public bool IncludeMold;
        public bool IncludeAerobic;
        public bool IncludeColiform;
        public bool IncludeEColi;
        public bool IncludeLactics;
        public bool IncludeSalmonella;
        public bool IncludeListeria;
        public bool IncludeColorAndAppearance;
        public bool IncludeForm;
        public bool IncludeFlavorAndOdor;

        public Template ()
        {
            if (Directory.Exists(Environment.GetFolderPath(Environment.SpecialFolder.Desktop) + "\\CoAs\\Templates") == false)
                Directory.CreateDirectory(Environment.GetFolderPath(Environment.SpecialFolder.Desktop) + "\\CoAs\\Templates");

            Menu = new Console.SelectionMenu(GetOptions(), "Templates:", "Please select a template");

            AssignOptionsFromFile(Menu.UserChoice);
        }
        /// <summary>
        /// Fetches names of available templates
        /// </summary>
        /// <returns></returns>
        private List<string> GetOptions()
        {
            string templateDirectory = Environment.GetFolderPath(Environment.SpecialFolder.Desktop) + "\\CoAs\\Templates";
            
            // A forever loop is triggered if no template files are found
            if(Directory.GetFiles(templateDirectory, "*.txt").Length == 0)
            {
                do
                {
                    Console.Util.WriteMessageInCenter("Could not find any templates in " + templateDirectory +
                        "  Press any key once templates have been added", ConsoleColor.Red);

                    System.Console.ReadKey();
                    System.Console.Write("\b \b ");

                } while (Directory.GetFiles(templateDirectory, "*.txt").Length == 0);

                Console.Util.RemoveMessageInCenter();
            }

            return KeepOnlyFileNames(Directory.GetFiles(templateDirectory, "*.txt")).ToList();
        }
        /// <summary>
        /// Reassigns array values from full file paths to only file names, excluding extensions
        /// </summary>
        /// <param name="targetArray"></param>
        /// <returns></returns>
        private string[] KeepOnlyFileNames(string[] targetArray)
        {
            List<string> tempStore;

            for (int i = 0; i < targetArray.Length; i++)
            {
                tempStore = targetArray[i].Split(new char[] { '.', '\\', '/' }).ToList();
                targetArray[i] = tempStore[tempStore.Count - 2];
            }

            return targetArray;
        }
        /// <summary>
        /// Parses the template file and assigns class variables accordingly
        /// </summary>
        /// <param name="templateName"></param>
        private void AssignOptionsFromFile(string templateName)
        {
            ContentCategories currentCategory = ContentCategories.None; // Default assignment, not used by any selection statements

            List<string> delimitedLine; // Each line consists of an option title/name, an equals sign, and the option choice

            Console.Util.WriteMessageInCenter("Loading " + Menu.UserChoice + " Template");

            foreach(string line in File.ReadLines(Environment.GetFolderPath(Environment.SpecialFolder.Desktop) + "\\CoAs\\Templates\\" + templateName + ".txt"))
            {
                delimitedLine = line.Split(new char[] { '=' }).ToList();

                // Trims strings of whitespace
                for(int i = 0; i < delimitedLine.Count; i++)
                {
                    delimitedLine[i] = delimitedLine[i].Trim();
                }
                // Whether the line contains a category title is determined first as it informs succeeding selection statements
                switch(line.ToLower())
                {
                    case "[algorithm]":
                        currentCategory = ContentCategories.Algorithm;
                        break;
                    case "[main content block]":
                        currentCategory = ContentCategories.MainContentBlock;
                        break;
                    default:
                        break;
                }
                // Sets SelectedAlgorithm, if applicable
                if(currentCategory == ContentCategories.Algorithm)
                {
                    // A switch statement is used to accomodate additions of template options
                    switch(delimitedLine[0].ToLower())
                    {
                        case "type":
                            switch(delimitedLine[1].ToLower())
                            {
                                case "resultsfromdateonwards":
                                    SelectedAlgorithm = Algorithm.ResultsFromDateOnwards;
                                    break;
                                case "standard":
                                    SelectedAlgorithm = Algorithm.Standard;
                                    break;
                                default:
                                    break;
                            }
                            break;
                        default:
                            break;
                    }
                }
                // Sets class bools for content inclusion
                else if(currentCategory == ContentCategories.MainContentBlock)
                {
                    switch (delimitedLine[0].ToLower())
                    {
                        case "customer name":
                            if (delimitedLine[1].ToLower() == "true")
                                IncludeCustomerName = true;
                            else
                                IncludeCustomerName = false;
                            break;
                        case "sales order":
                            if (delimitedLine[1].ToLower() == "true")
                                IncludeSalesOrder = true;
                            else
                                IncludeSalesOrder = false;
                            break;
                        case "purchase order":
                            if (delimitedLine[1].ToLower() == "true")
                                IncludePurchaseOrder = true;
                            else
                                IncludePurchaseOrder = false;
                            break;
                        case "generation date":
                            if (delimitedLine[1].ToLower() == "true")
                                IncludeGenerationDate = true;
                            else
                                IncludeGenerationDate = false;
                            break;
                        case "product name":
                            if (delimitedLine[1].ToLower() == "true")
                                IncludeProductName = true;
                            else
                                IncludeProductName = false;
                            break;
                        case "recipe/item":
                            if (delimitedLine[1].ToLower() == "true")
                                IncludeRecipeAndItem = true;
                            else
                                IncludeRecipeAndItem = false;
                            break;
                        case "lot code":
                            if (delimitedLine[1].ToLower() == "true")
                                IncludeLotCode = true;
                            else
                                IncludeLotCode = false;
                            break;
                        case "batch":
                            if (delimitedLine[1].ToLower() == "true")
                                IncludeBatch = true;
                            else
                                IncludeBatch = false;
                            break;
                        case "best by date":
                            if (delimitedLine[1].ToLower() == "true")
                                IncludeBestByDate = true;
                            else
                                IncludeBestByDate = false;
                            break;
                        case "manufacturing site":
                            if (delimitedLine[1].ToLower() == "true")
                                IncludeManufacturingSite = true;
                            else
                                IncludeManufacturingSite = false;
                            break;
                        case "manufacturing date":
                            if (delimitedLine[1].ToLower() == "true")
                                IncludeManufacturingDate = true;
                            else
                                IncludeManufacturingDate = false;
                            break;
                        case "acidity":
                            if (delimitedLine[1].ToLower() == "true")
                                IncludeAcidity = true;
                            else
                                IncludeAcidity = false;
                            break;
                        case "ph":
                            if (delimitedLine[1].ToLower() == "true")
                                IncludepH = true;
                            else
                                IncludepH = false;
                            break;
                        case "viscosity cm":
                            if (delimitedLine[1].ToLower() == "true")
                                IncludeViscosityCM = true;
                            else
                                IncludeViscosityCM = false;
                            break;
                        case "viscosity cps":
                            if (delimitedLine[1].ToLower() == "true")
                                IncludeViscosityCPS = true;
                            else
                                IncludeViscosityCPS = false;
                            break;
                        case "water activity":
                            if (delimitedLine[1].ToLower() == "true")
                                IncludeWaterActivity = true;
                            else
                                IncludeWaterActivity = false;
                            break;
                        case "brix slurry":
                            if (delimitedLine[1].ToLower() == "true")
                                IncludeBrixSlurry = true;
                            else
                                IncludeBrixSlurry = false;
                            break;
                        case "yeast":
                            if (delimitedLine[1].ToLower() == "true")
                                IncludeYeast = true;
                            else
                                IncludeYeast = false;
                            break;
                        case "mold":
                            if (delimitedLine[1].ToLower() == "true")
                                IncludeMold = true;
                            else
                                IncludeMold = false;
                            break;
                        case "aerobic":
                            if (delimitedLine[1].ToLower() == "true")
                                IncludeAerobic = true;
                            else
                                IncludeAerobic = false;
                            break;
                        case "coliform":
                            if (delimitedLine[1].ToLower() == "true")
                                IncludeColiform = true;
                            else
                                IncludeColiform = false;
                            break;
                        case "ecoli":
                            if (delimitedLine[1].ToLower() == "true")
                                IncludeEColi = true;
                            else
                                IncludeEColi = false;
                            break;
                        case "lactics":
                            if (delimitedLine[1].ToLower() == "true")
                                IncludeLactics = true;
                            else
                                IncludeLactics = false;
                            break;
                        case "salmonella":
                            if (delimitedLine[1].ToLower() == "true")
                                IncludeSalmonella = true;
                            else
                                IncludeSalmonella = false;
                            break;
                        case "listeria":
                            if (delimitedLine[1].ToLower() == "true")
                                IncludeListeria = true;
                            else
                                IncludeListeria = false;
                            break;
                        case "color/appearance":
                            if (delimitedLine[1].ToLower() == "true")
                                IncludeColorAndAppearance = true;
                            else
                                IncludeColorAndAppearance = false;
                            break;
                        case "form":
                            if (delimitedLine[1].ToLower() == "true")
                                IncludeForm = true;
                            else
                                IncludeForm = false;
                            break;
                        case "flavor/odor":
                            if (delimitedLine[1].ToLower() == "true")
                                IncludeFlavorAndOdor = true;
                            else
                                IncludeFlavorAndOdor = false;
                            break;
                        default:
                            break;
                    }
                }
                
            }

            Console.Util.RemoveMessageInCenter(); // Removes message written just before foreach loop was initiated
        }

    }
}
