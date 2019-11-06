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
            Unassigned, CustomerName, SalesOrder, PurchaseOrder, GenerationDate, ProductName, RecipeAndItem, LotCode, Batch,
            BestByDate, ManufacturingSite, ManufacturingDate, Acidity, pH, ViscosityCM, ViscosityCPS, WaterActivity, BrixSlurry, Yeast, Mold,
            Aerobic, Coliform, EColi, Lactics, Salmonella, Listeria, ColorAndAppearance, Form, FlavorAndOdor
        }

        // Enum variables
        public Algorithm SelectedAlgorithm;

        // Lists
        public List<Templates.CustomSearch> CustomSearches;
        public List<Templates.CustomFilter> CustomFilters;

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

        public Console.SelectionMenu Menu;

        public Template ()
        {
            SetInclusionBoolsAsFalse();

            if (Directory.Exists(Environment.GetFolderPath(Environment.SpecialFolder.Desktop) + "\\CoAs\\Templates") == false)
                Directory.CreateDirectory(Environment.GetFolderPath(Environment.SpecialFolder.Desktop) + "\\CoAs\\Templates");

            CustomSearches = new List<Templates.CustomSearch>();

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

                    case "[filter coa]":
                        currentCategory = ContentCategories.FilterCOA;
                        break;

                    case "[include content items]":
                        currentCategory = ContentCategories.IncludeContentItems;
                        break;

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
                                    if (new Console.SelectionMenu(optionsForInvalidItem, "", promptForInvalidItem).UserChoice == "Exit application")
                                        Environment.Exit(0);
                                    break;
                            }
                            break;

                        case "data-group to search":
                            {
                                bool assignedDataGroup = false;

                                switch (delimitedLine[1].ToLower())
                                {
                                    case "micro":

                                        foreach(Templates.CustomSearch search in CustomSearches)
                                        {
                                            if(search.DataGroup != Templates.CustomSearch.DataGroupToSearch.Micro ||
                                                search.DataGroup != Templates.CustomSearch.DataGroupToSearch.Unassigned)
                                            {
                                                Console.Util.WriteMessageInCenter("Program cannot proceed with mixed data-groups in algorithm information.  " +
                                                    "  Press any key to exit application.", ConsoleColor.Red);
                                                System.Console.ReadKey();
                                                Environment.Exit(0);
                                            }
                                        }

                                        do // Find customSearch with unassigned datagroup, if can't, make a new customSearch and try again
                                        {
                                            foreach (Templates.CustomSearch customSearch in CustomSearches)
                                            {
                                                if (customSearch.DataGroup == Templates.CustomSearch.DataGroupToSearch.Unassigned)
                                                {
                                                    customSearch.DataGroup = Templates.CustomSearch.DataGroupToSearch.Micro;
                                                    assignedDataGroup = true;
                                                }
                                            }
                                            if (assignedDataGroup == false)
                                            {
                                                CustomSearches.Add(new Templates.CustomSearch());
                                            }
                                        } while (assignedDataGroup == false);
                                        break;
                                    default:
                                        promptForInvalidItem = delimitedLine[1] + " is not a valid algorithm data-group to search.";
                                        if (new Console.SelectionMenu(optionsForInvalidItem, "", promptForInvalidItem).UserChoice == "Exit application")
                                            Environment.Exit(0);
                                        break;
                                }
                            }
                            break;

                        case "where value":
                            bool assignedSearchCriteria = false;

                            do // Find customSearch with unassigned datagroup, if can't, make a new customSearch and try again
                            {
                                foreach (Templates.CustomSearch customSearch in CustomSearches)
                                {
                                    if (customSearch.SearchCriteria == string.Empty)
                                    {
                                        customSearch.SearchCriteria = delimitedLine[1];
                                        assignedSearchCriteria = true;
                                    }
                                }
                                if (assignedSearchCriteria == false)
                                {
                                    CustomSearches.Add(new Templates.CustomSearch());
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
                                        foreach (Templates.CustomSearch customSearch in CustomSearches)
                                        {
                                            if (customSearch.SearchColumnOffset == -1) // -1 is the default value
                                            {
                                                customSearch.SearchColumnOffset = searchColumnValue;
                                                assignedSearchColumn = true;
                                            }
                                        }
                                        if (assignedSearchColumn == false)
                                        {
                                            CustomSearches.Add(new Templates.CustomSearch());
                                        }
                                    } while (assignedSearchColumn == false);
                                }
                                else
                                {
                                    promptForInvalidItem = delimitedLine[1] + " is not a valid number for algorithm search column";
                                    if (new Console.SelectionMenu(optionsForInvalidItem, "", promptForInvalidItem).UserChoice == "Exit application")
                                        Environment.Exit(0);
                                }
                            }
                            break;

                        case "":
                            break;

                        default:
                            promptForInvalidItem = delimitedLine[1] + " is not a valid algorithm item.";
                            new Console.SelectionMenu(optionsForInvalidItem, "", promptForInvalidItem);
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
                                do // Find customFilter with unassigned Filter, if can't, make a new customFilter and try again
                                {
                                    foreach (Templates.CustomFilter customFilter in CustomFilters)
                                    {
                                        if (customFilter.Filter == Templates.CustomFilter.FilterType.Unassigned)
                                        {
                                            forCustomFilter = CustomFilters.Count - 1;

                                            if (delimitedLine[1].ToLower() == "whitelist")
                                                customFilter.Filter = Templates.CustomFilter.FilterType.Whitelist;
                                            else
                                                customFilter.Filter = Templates.CustomFilter.FilterType.Blacklist;
                                            
                                            assignedFilter = true;
                                        }
                                    }
                                    if (assignedFilter == false)
                                        CustomSearches.Add(new Templates.CustomSearch());
                                    
                                } while (assignedFilter == false);
                            }
                            break;
                        case "filter by item":
                            {
                                bool assignedContentItem = false;
                                do // Find customFilter with unassigned ContentItem, if can't, make a new customFilter and try again
                                {
                                    foreach(Templates.CustomFilter customFilter in CustomFilters)
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
                                                case "batch":
                                                    customFilter.ContentItem = ContentItems.Batch;
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
                                                    
                                                    break;
                                                // CustomerName, SalesOrder, PurchaseOrder, GenerationDate, ProductName, RecipeAndItem, LotCode, Batch,
                                                // BestByDate, ManufacturingSite, ManufacturingDate, Acidity, pH, ViscosityCM, ViscosityCPS, WaterActivity, BrixSlurry, Yeast, Mold,
                                                // Aerobic, Coliform, EColi, Lactics, Salmonella, Listeria, ColorAndAppearance, Form, FlavorAndOdor
                                                case "":
                                                    break;
                                                default:
                                                    promptForInvalidItem = delimitedLine[1] + " is not a valid content item";
                                                    new Console.SelectionMenu(optionsForInvalidItem, "", promptForInvalidItem);
                                                    break;
                                            }

                                        if (assignedContentItem == false)
                                            CustomFilters.Add(new Templates.CustomFilter());
                                    }
                                } while (assignedContentItem == false);
                            }
                            break;
                        case "where item":
                            {
                                if(CustomFilters.Count == 0)
                                {
                                    CustomFilters.Add(new Templates.CustomFilter());
                                }

                                CustomFilters[forCustomFilter].Criteria.Add(delimitedLine[1]);
                            }
                            break;
                        case "":
                            break;
                        default:
                            promptForInvalidItem = delimitedLine[0] + " is not a valid filter item.";
                            new Console.SelectionMenu(optionsForInvalidItem, "", promptForInvalidItem);
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
                        case "batch":
                            if (delimitedLine[1].ToLower() == "yes")
                                IncludeBatch = true;
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
                            new Console.SelectionMenu(optionsForInvalidItem, "", promptForInvalidItem);
                            break;
                    }
                
                
            }

            Console.Util.RemoveMessageInCenter(); // Removes message written just before foreach loop was initiated
        }
        /// <summary>
        /// Sets all content inclusion bools to their intended default, false
        /// </summary>
        private void SetInclusionBoolsAsFalse()
        {
            IncludeCustomerName = false;
            IncludeSalesOrder = false;
            IncludePurchaseOrder = false;
            IncludeGenerationDate = false;
            IncludeProductName = false;
            IncludeRecipeAndItem = false;
            IncludeLotCode = false;
            IncludeBatch = false;
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
