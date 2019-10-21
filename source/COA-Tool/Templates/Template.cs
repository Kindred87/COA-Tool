using System;
using System.Collections.Generic;
using System.Text;
using System.IO;


namespace CoA_Tool
{
    class Template
    {
        public enum SearchTargets { Product, RecipeAndItem,  LotCode, ManufactureSite, Acidity, pH, ViscosityCPS, 
            Yeast, Mold, Aerobic, Coliform, Lactics}
        public List<SearchTargets> TemplateContents = new List<SearchTargets>();
        public Template ()
        {

        }
        /// <summary>
        /// Fetches array of template options
        /// </summary>
        /// <returns></returns>
        public string[] GetOptions()
        {
            return Directory.GetFiles(Directory.GetCurrentDirectory() + "/Templates", "*.txt");
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
