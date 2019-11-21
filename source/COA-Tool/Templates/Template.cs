using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.IO;
using CoA_Tool;
using CoA_Tool.Utility;


namespace CoA_Tool.Templates
{
    /// <summary>
    /// Contains information on how the Excel document is populated
    /// </summary>
    class Template
    {
        // Enums
        /// <summary>
        /// General methods for generating unique varieties of CoAs
        /// </summary>
        public enum Algorithm { Standard, FromDateOnwards}
        private enum ContentCategories { None, Algorithm, FilterCOA, IncludeContentItems}
        /// <summary>
        /// Represents the different items that can be included in a CoA document
        /// </summary>
        public enum ContentItems
        {
            Unassigned, CustomerName, SalesOrder, PurchaseOrder, GenerationDate, ProductName, RecipeAndItem, LotCode, BatchFromMicro, BatchFromDressing,
            BestByDate, ManufacturingSite, ManufacturingDate, Acidity, pH, ViscosityCM, ViscosityCPS, WaterActivity, BrixSlurry, Yeast, Mold,
            Aerobic, Coliform, EColi, Lactics, Salmonella, Listeria, ColorAndAppearance, Form, FlavorAndOdor
        }

        // Enum variables
        public Algorithm SelectedAlgorithm;

        // Lists
        public List<CustomSearch> CustomSearches;
        public List<CustomFilter> CustomFilters;

        // Bools
        public bool IncludeCustomerName;
        public bool IncludeSalesOrder;
        public bool IncludePurchaseOrder;
        public bool IncludeGenerationDate;
        public bool IncludeProductName;
        public bool IncludeRecipeAndItem;
        public bool IncludeLotCode;
        public bool IncludeBatchFromMicro;
        public bool IncludeBatchFromDressing;
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

        // Objects
        public ConsoleInteraction.SelectionMenu Menu;

        // Strings
        /// <summary>
        /// Synonymous with the template's file name
        /// </summary>
        public string CustomerName
        {
            get
            {
                return Menu.UserChoice;
            }
        }

        public Template ()
        {
            SetInclusionBoolsToFalse();

            if (Directory.Exists(Environment.GetFolderPath(Environment.SpecialFolder.Desktop) + "\\CoAs\\Templates") == false)
                Directory.CreateDirectory(Environment.GetFolderPath(Environment.SpecialFolder.Desktop) + "\\CoAs\\Templates");

            CustomSearches = new List<CustomSearch>();
            CustomFilters = new List<CustomFilter>();

            Menu = new ConsoleInteraction.SelectionMenu(GetOptions(), "Templates:", "Please select a template");

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
                    ConsoleOps.WriteMessageInCenter("Could not find any templates in " + templateDirectory +
                        "  Press any key once templates have been added", ConsoleColor.Red);

                    System.Console.ReadKey();
                    System.Console.Write("\b \b ");

                } while (Directory.GetFiles(templateDirectory, "*.txt").Length == 0);

                ConsoleOps.RemoveMessageInCenter();
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

            ConsoleOps.WriteMessageInCenter("Loading " + Menu.UserChoice + " Template");

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
                        continue;

                    case "[filter coa]":
                        currentCategory = ContentCategories.FilterCOA;
                        continue;

                    case "[include content items]":
                        currentCategory = ContentCategories.IncludeContentItems;
                        continue;

                    default:
                        break;
                }

                // These are used with Console.SelectionMenu in switch defaults to notify...
                // the user of invalid items in the template file
                string promptForInvalidItem;
                List<string> optionsForInvalidItem = new List<string>();
                optionsForInvalidItem.Add("Continue regardless");
                optionsForInvalidItem.Add("Exit application");

                // Sets SelectedAlgorithm, if applicable
                if(currentCategory == ContentCategories.Algorithm)
                    switch(delimitedLine[0].ToLower()) // A switch statement is used to accomodate additions of template options
                    {
                        case "type":
                            switch(delimitedLine[1].ToLower())
                            {
                                case "fromdateonwards":
                                    SelectedAlgorithm = Algorithm.FromDateOnwards;
                                    break;

                                case "standard":
                                    SelectedAlgorithm = Algorithm.Standard;
                                    break;

                                default:
                                    promptForInvalidItem = delimitedLine[1] + " is not a valid algorithm type.";
                                    if (new ConsoleInteraction.SelectionMenu(optionsForInvalidItem, "", promptForInvalidItem).UserChoice == "Exit application")
                                    {
                                        Environment.Exit(0);
                                    }
                                    break;
                            }
                            break;

                        case "search in":
                            {
                                bool assignedDataGroup = false;

                                switch (delimitedLine[1].ToLower())
                                {
                                    case "micro":

                                        foreach(CustomSearch search in CustomSearches)
                                        {
                                            if(search.DataGroup != CustomSearch.DataGroupToSearch.Micro ||
                                                search.DataGroup != CustomSearch.DataGroupToSearch.Unassigned)
                                            {
                                                //TODO: Replace with acknowledgment menu
                                                ConsoleOps.WriteMessageInCenter("Program cannot proceed with mixed \"search in\" targets in algorithm information.  " +
                                                    "  Press any key to exit application.", ConsoleColor.Red);
                                                System.Console.ReadKey();
                                                Environment.Exit(0);
                                            }
                                        }

                                        do // Find customSearch with unassigned datagroup, if can't, make a new customSearch and try again
                                        {
                                            foreach (CustomSearch customSearch in CustomSearches)
                                            {
                                                if (customSearch.DataGroup == CustomSearch.DataGroupToSearch.Unassigned)
                                                {
                                                    customSearch.DataGroup = CustomSearch.DataGroupToSearch.Micro;
                                                    assignedDataGroup = true;
                                                }
                                            }
                                            if (assignedDataGroup == false)
                                            {
                                                CustomSearches.Add(new CustomSearch());
                                            }
                                        } while (assignedDataGroup == false);
                                        break;

                                    case "":
                                        break;

                                    default:
                                        promptForInvalidItem = delimitedLine[1] + " is not a valid item to search in.";
                                        if (new ConsoleInteraction.SelectionMenu(optionsForInvalidItem, "", promptForInvalidItem).UserChoice == "Exit application")
                                        {
                                            Environment.Exit(0);
                                        }
                                        break;
                                }
                            }
                            break;

                        case "where value":
                            bool assignedSearchCriteria = false;

                            do // Find customSearch with unassigned datagroup, if can't, make a new customSearch and try again
                            {
                                foreach (CustomSearch customSearch in CustomSearches)
                                {
                                    if (customSearch.SearchCriteria == string.Empty)
                                    {
                                        customSearch.SearchCriteria = delimitedLine[1];
                                        assignedSearchCriteria = true;
                                    }
                                }
                                if (assignedSearchCriteria == false)
                                {
                                    CustomSearches.Add(new CustomSearch());
                                }
                            } while (assignedSearchCriteria == false);
                            break;

                        case "in column":
                            {
                                bool assignedSearchColumn = false;

                                if (Int32.TryParse(delimitedLine[1], out int searchColumnValue))
                                {
                                    do // Find customSearch with unassigned datagroup, if can't, make a new customSearch and try again
                                    {
                                        foreach (CustomSearch customSearch in CustomSearches)
                                        {
                                            if (customSearch.SearchColumnOffset == -1) // -1 is the default value
                                            {
                                                customSearch.SearchColumnOffset = searchColumnValue;
                                                assignedSearchColumn = true;
                                            }
                                        }
                                        if (assignedSearchColumn == false)
                                        {
                                            CustomSearches.Add(new CustomSearch());
                                        }
                                    } while (assignedSearchColumn == false);
                                }
                                else if(delimitedLine[1] != "")
                                {
                                    promptForInvalidItem = delimitedLine[1] + " is not a valid number for algorithm search column";
                                    if (new ConsoleInteraction.SelectionMenu(optionsForInvalidItem, "", promptForInvalidItem).UserChoice == "Exit application")
                                    {
                                        Environment.Exit(0);
                                    }
                                        
                                }
                            }
                            break;

                        case "":
                            break;

                        default:
                            promptForInvalidItem = "\"" + delimitedLine[0] + "\"" +" is not a valid algorithm item.";
                            if(new ConsoleInteraction.SelectionMenu(optionsForInvalidItem, "", promptForInvalidItem).UserChoice == "Exit application")
                            {
                                Environment.Exit(0);
                            }
                            break;
                    }
                
                else if(currentCategory == ContentCategories.FilterCOA)
                {
                    // Used to track which filter to apply criteria to, assigned with filter in/out and content item
                    int forCustomFilter = 0;

                    switch (delimitedLine[0].ToLower())
                    {
                        case "whitelist or blacklist":
                            {
                                bool assignedFilter = false;
                                do // Find customFilter with unassigned FilterType, if can't, make a new customFilter and try again
                                {
                                    foreach (CustomFilter customFilter in CustomFilters)
                                    {
                                        if (customFilter.FilterType == CustomFilter.FilterTypes.Unassigned)
                                        {
                                            forCustomFilter = CustomFilters.Count - 1;

                                            if (delimitedLine[1].ToLower() == "whitelist")
                                                customFilter.FilterType = CustomFilter.FilterTypes.Whitelist;
                                            else
                                                customFilter.FilterType = CustomFilter.FilterTypes.Blacklist;
                                            
                                            assignedFilter = true;
                                        }
                                    }
                                    if (assignedFilter == false)
                                        CustomFilters.Add(new CustomFilter());
                                    
                                } while (assignedFilter == false);
                            }
                            break;
                        case "filter by item":
                            {
                                bool assignedContentItem = false;
                                do // Find customFilter with unassigned ContentItem, if can't, make a new customFilter and try again
                                {
                                    foreach(CustomFilter customFilter in CustomFilters)
                                    {
                                        forCustomFilter = CustomFilters.Count - 1;

                                        if (customFilter.ContentItem == ContentItems.Unassigned)
                                            switch(delimitedLine[1].ToLower())
                                            {
                                                case "customer name":
                                                    customFilter.ContentItem = ContentItems.CustomerName;
                                                    assignedContentItem = true;
                                                    break;
                                                case "sales order":
                                                    customFilter.ContentItem = ContentItems.SalesOrder;
                                                    assignedContentItem = true;
                                                    break;
                                                case "purchase order":
                                                    customFilter.ContentItem = ContentItems.PurchaseOrder;
                                                    assignedContentItem = true;
                                                    break;
                                                case "generation date":
                                                    customFilter.ContentItem = ContentItems.GenerationDate;
                                                    assignedContentItem = true;
                                                    break;
                                                case "product name":
                                                    customFilter.ContentItem = ContentItems.ProductName;
                                                    assignedContentItem = true;
                                                    break;
                                                case "recipe/item":
                                                    customFilter.ContentItem = ContentItems.RecipeAndItem;
                                                    assignedContentItem = true;
                                                    break;
                                                case "lot code":
                                                    customFilter.ContentItem = ContentItems.LotCode;
                                                    assignedContentItem = true;
                                                    break;
                                                case "batch from micro":
                                                    customFilter.ContentItem = ContentItems.BatchFromMicro;
                                                    assignedContentItem = true;
                                                    break;
                                                case "batch from dressing":
                                                    customFilter.ContentItem = ContentItems.BatchFromDressing;
                                                    assignedContentItem = true;
                                                    break;
                                                case "best by date":
                                                    customFilter.ContentItem = ContentItems.BestByDate;
                                                    assignedContentItem = true;
                                                    break;
                                                case "manufacturing site":
                                                    customFilter.ContentItem = ContentItems.ManufacturingSite;
                                                    assignedContentItem = true;
                                                    break;
                                                case "manufacturing date":
                                                    customFilter.ContentItem = ContentItems.ManufacturingDate;
                                                    assignedContentItem = true;
                                                    break;
                                                case "acidity":
                                                    customFilter.ContentItem = ContentItems.Acidity;
                                                    assignedContentItem = true;
                                                    break;
                                                case "ph":
                                                    customFilter.ContentItem = ContentItems.pH;
                                                    assignedContentItem = true;
                                                    break;
                                                case "viscosity cm":
                                                    customFilter.ContentItem = ContentItems.ViscosityCM;
                                                    assignedContentItem = true;
                                                    break;
                                                case "viscosity cps":
                                                    customFilter.ContentItem = ContentItems.ViscosityCPS;
                                                    assignedContentItem = true;
                                                    break;
                                                case "water activity":
                                                    customFilter.ContentItem = ContentItems.WaterActivity;
                                                    assignedContentItem = true;
                                                    break;
                                                case "brix slurry":
                                                    customFilter.ContentItem = ContentItems.BrixSlurry;
                                                    assignedContentItem = true;
                                                    break;
                                                case "yeast":
                                                    customFilter.ContentItem = ContentItems.Yeast;
                                                    assignedContentItem = true;
                                                    break;
                                                case "mold":
                                                    customFilter.ContentItem = ContentItems.Mold;
                                                    assignedContentItem = true;
                                                    break;
                                                case "aerobic":
                                                    customFilter.ContentItem = ContentItems.Aerobic;
                                                    assignedContentItem = true;
                                                    break;
                                                case "coliform":
                                                    customFilter.ContentItem = ContentItems.Coliform;
                                                    assignedContentItem = true;
                                                    break;
                                                case "e. coli":
                                                    customFilter.ContentItem = ContentItems.EColi;
                                                    assignedContentItem = true;
                                                    break;
                                                case "lactics":
                                                    customFilter.ContentItem = ContentItems.Lactics;
                                                    assignedContentItem = true;
                                                    break;
                                                case "salmonella":
                                                    customFilter.ContentItem = ContentItems.Salmonella;
                                                    assignedContentItem = true;
                                                    break;
                                                case "listeria":
                                                    customFilter.ContentItem = ContentItems.Listeria;
                                                    assignedContentItem = true;
                                                    break;
                                                case "color/appearance":
                                                    customFilter.ContentItem = ContentItems.ColorAndAppearance;
                                                    assignedContentItem = true;
                                                    break;
                                                case "form":
                                                    customFilter.ContentItem = ContentItems.Form;
                                                    assignedContentItem = true;
                                                    break;
                                                case "flavor/odor":
                                                    customFilter.ContentItem = ContentItems.FlavorAndOdor;
                                                    assignedContentItem = true;
                                                    break;
                                                case "":
                                                    assignedContentItem = true; // Having no value is considered valid
                                                    break;
                                                default:
                                                    promptForInvalidItem = delimitedLine[1] + " is not a valid content item";
                                                    if(new ConsoleInteraction.SelectionMenu(optionsForInvalidItem, "", promptForInvalidItem).UserChoice == "Exit application")
                                                    {
                                                        Environment.Exit(0);
                                                    }
                                                    break;
                                            }
                                    }

                                    if (assignedContentItem == false)
                                    {
                                        CustomFilters.Add(new CustomFilter());
                                    }

                                } while (assignedContentItem == false);
                            }
                            break;
                        case "where item":
                            {
                                if(CustomFilters.Count == 0)
                                {
                                    CustomFilters.Add(new CustomFilter());
                                }

                                CustomFilters[forCustomFilter].Criteria.Add(delimitedLine[1]);
                            }
                            break;
                        case "":
                            break;
                        default:
                            promptForInvalidItem = delimitedLine[0] + " is not a valid filter item.";
                            if(new ConsoleInteraction.SelectionMenu(optionsForInvalidItem, "", promptForInvalidItem).UserChoice == "Exit application")
                            {
                                Environment.Exit(0);
                            }
                            break;
                    }
                }
                // Sets class bools for content inclusion
                else if(currentCategory == ContentCategories.IncludeContentItems)
                    switch (delimitedLine[0].ToLower())
                    {
                        case "customer name":
                            if (delimitedLine[1].ToLower() == "yes")
                                IncludeCustomerName = true;
                            break;
                        case "sales order":
                            if (delimitedLine[1].ToLower() == "yes")
                                IncludeSalesOrder = true;
                            break;
                        case "purchase order":
                            if (delimitedLine[1].ToLower() == "yes")
                                IncludePurchaseOrder = true;
                            break;
                        case "generation date":
                            if (delimitedLine[1].ToLower() == "yes")
                                IncludeGenerationDate = true;
                            break;
                        case "product name":
                            if (delimitedLine[1].ToLower() == "yes")
                                IncludeProductName = true;
                            break;
                        case "recipe/item":
                            if (delimitedLine[1].ToLower() == "yes")
                                IncludeRecipeAndItem = true;
                            break;
                        case "lot code":
                            if (delimitedLine[1].ToLower() == "yes")
                                IncludeLotCode = true;
                            break;
                        case "batch from micro":
                            if (delimitedLine[1].ToLower() == "yes")
                                IncludeBatchFromMicro = true;
                            break;
                        case "batch from dressing":
                            if (delimitedLine[1].ToLower() == "yes")
                                IncludeBatchFromDressing = true;
                            break;
                        case "best by date":
                            if (delimitedLine[1].ToLower() == "yes")
                                IncludeBestByDate = true;
                            break;
                        case "manufacturing site":
                            if (delimitedLine[1].ToLower() == "yes")
                                IncludeManufacturingSite = true;
                            break;
                        case "manufacturing date":
                            if (delimitedLine[1].ToLower() == "yes")
                                IncludeManufacturingDate = true;
                            break;
                        case "acidity":
                            if (delimitedLine[1].ToLower() == "yes")
                                IncludeAcidity = true;
                            break;
                        case "ph":
                            if (delimitedLine[1].ToLower() == "yes")
                                IncludepH = true;
                            break;
                        case "viscosity cm":
                            if (delimitedLine[1].ToLower() == "yes")
                                IncludeViscosityCM = true;
                            break;
                        case "viscosity cps":
                            if (delimitedLine[1].ToLower() == "yes")
                                IncludeViscosityCPS = true;
                            break;
                        case "water activity":
                            if (delimitedLine[1].ToLower() == "yes")
                                IncludeWaterActivity = true;
                            break;
                        case "brix slurry":
                            if (delimitedLine[1].ToLower() == "yes")
                                IncludeBrixSlurry = true;
                            break;
                        case "yeast":
                            if (delimitedLine[1].ToLower() == "yes")
                                IncludeYeast = true;
                            break;
                        case "mold":
                            if (delimitedLine[1].ToLower() == "yes")
                                IncludeMold = true;
                            break;
                        case "aerobic":
                            if (delimitedLine[1].ToLower() == "yes")
                                IncludeAerobic = true;
                            break;
                        case "coliform":
                            if (delimitedLine[1].ToLower() == "yes")
                                IncludeColiform = true;
                            break;
                        case "e. coli":
                            if (delimitedLine[1].ToLower() == "yes")
                                IncludeEColi = true;
                            break;
                        case "lactics":
                            if (delimitedLine[1].ToLower() == "yes")
                                IncludeLactics = true;
                            break;
                        case "salmonella":
                            if (delimitedLine[1].ToLower() == "yes")
                                IncludeSalmonella = true;
                            break;
                        case "listeria":
                            if (delimitedLine[1].ToLower() == "yes")
                                IncludeListeria = true;
                            break;
                        case "color/appearance":
                            if (delimitedLine[1].ToLower() == "yes")
                                IncludeColorAndAppearance = true;
                            break;
                        case "form":
                            if (delimitedLine[1].ToLower() == "yes")
                                IncludeForm = true;
                            break;
                        case "flavor/odor":
                            if (delimitedLine[1].ToLower() == "yes")
                                IncludeFlavorAndOdor = true;
                            break;
                        case "":
                            break;
                        default:
                            promptForInvalidItem = delimitedLine[0] + " is not a valid content item.";
                            if(new ConsoleInteraction.SelectionMenu(optionsForInvalidItem, "", promptForInvalidItem).UserChoice == "Exit application")
                            {
                                Environment.Exit(0);
                            }
                            break;
                    }
                
                
            }

            ConsoleOps.RemoveMessageInCenter(); // Removes message written just before foreach loop was initiated
        }
        /// <summary>
        /// Sets all content inclusion bools to their intended default, false
        /// </summary>
        private void SetInclusionBoolsToFalse()
        {
            IncludeCustomerName = false;
            IncludeSalesOrder = false;
            IncludePurchaseOrder = false;
            IncludeGenerationDate = false;
            IncludeProductName = false;
            IncludeRecipeAndItem = false;
            IncludeLotCode = false;
            IncludeBatchFromMicro = false;
            IncludeBatchFromDressing = false;
            IncludeBestByDate = false;
            IncludeManufacturingDate = false;
            IncludeManufacturingSite = false;
            IncludeAcidity = false;
            IncludepH = false;
            IncludeViscosityCM = false;
            IncludeViscosityCPS = false;
            IncludeWaterActivity = false;
            IncludeBrixSlurry = false;
            IncludeYeast = false;
            IncludeMold = false;
            IncludeAerobic = false;
            IncludeColiform = false;
            IncludeEColi = false;
            IncludeLactics = false;
            IncludeSalmonella = false;
            IncludeListeria = false;
            IncludeColorAndAppearance = false;
            IncludeForm = false;
            IncludeFlavorAndOdor = false;
        }

    }
}
