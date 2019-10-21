using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.IO;
using CoA_Tool;


namespace CoA_Tool
{
    class Template
    {
        public enum SearchTargets { Product, RecipeAndItem,  LotCode, ManufactureSite, Acidity, pH, ViscosityCPS, 
            Yeast, Mold, Aerobic, Coliform, Lactics}

        public List<SearchTargets> TemplateContents;

        private Templates.SelectionMenu Menu;

        public Template ()
        {
            // Almost entirely for development purposes
            if (Directory.Exists(Directory.GetCurrentDirectory() + "\\Templates") == false)
                Directory.CreateDirectory(Directory.GetCurrentDirectory() + "\\Templates");

            TemplateContents = new List<SearchTargets>();

            Menu = new Templates.SelectionMenu(GetOptions());
        }
        /// <summary>
        /// Fetches array of template options
        /// </summary>
        /// <returns></returns>
        private string[] GetOptions()
        {
            string templateDirectory = Directory.GetCurrentDirectory() + "\\Templates";
            
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

            string[] options = KeepOnlyFileNames(Directory.GetFiles(templateDirectory, "*.txt"));

            return options;
        }
        /// <summary>
        /// Reassigns array values from full file paths to ony file names, excluding extensions
        /// </summary>
        /// <param name="targetArray"></param>
        /// <returns></returns>
        private string[] KeepOnlyFileNames(string[] targetArray)
        {
            List<string> tempStore = new List<string>();

            for (int i = 0; i < targetArray.Length; i++)
            {
                tempStore = targetArray[i].Split(new char[] { '.', '\\', '/' }).ToList();
                targetArray[i] = tempStore[tempStore.Count - 2];
            }

            return targetArray;
        }
        /// <summary>
        /// Parses a template via name and populates TemplateContents accordingly
        /// </summary>
        /// <param name="template"></param>
        public void PopulateContents(string template)
        {
            foreach(string line in File.ReadLines(Directory.GetCurrentDirectory() + "/Templates/" + template + ".txt"))
            {
                switch(line)
                {
                    case "Product":
                        TemplateContents.Add(SearchTargets.Product);
                        break;
                    case "Recipe/Item":
                        TemplateContents.Add(SearchTargets.RecipeAndItem);
                        break;
                    case "Lot Code":
                        TemplateContents.Add(SearchTargets.LotCode);
                        break;
                    case "Manufacturing Site":
                        TemplateContents.Add(SearchTargets.ManufactureSite);
                        break;
                    case "Acidity":
                        TemplateContents.Add(SearchTargets.Acidity);
                        break;
                    case "pH":
                        TemplateContents.Add(SearchTargets.pH);
                        break;
                    case "Visocity cps":
                        TemplateContents.Add(SearchTargets.ViscosityCPS);
                        break;
                    case "Yeast":
                        TemplateContents.Add(SearchTargets.Yeast);
                        break;
                    case "Mold":
                        TemplateContents.Add(SearchTargets.Mold);
                        break;
                    case "Aerobic":
                        TemplateContents.Add(SearchTargets.Aerobic);
                        break;
                    case "Coliform":
                        TemplateContents.Add(SearchTargets.Coliform);
                        break;
                    case "Lactics":
                        TemplateContents.Add(SearchTargets.Lactics);
                        break;
                    default:
                        break;
                }
            }
        }

    }
}
