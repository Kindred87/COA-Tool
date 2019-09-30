using System;
using System.Collections.Generic;
using System.Text;
using System.IO;
using System.Linq;

namespace CoA_Tool.CSV
{
    class Tableau
    {
        public List<string> Paths;
        public List<List<List<string>>> FileContents;
        public Tableau()
        {
            FileContents = LoadFiles(OnlyValidPathsFrom(GetCSVFilePaths()));
        }
        /// <summary>
        /// Fetches list of csv file paths from the user's download folder
        /// </summary>
        /// <returns></returns>
        private string[] GetCSVFilePaths()
        {
            string[] directory = Directory.GetCurrentDirectory().Split(new char[] { '/', '\\' }, StringSplitOptions.None);
            string userFolder = directory[0] + "/" + directory[1] + "/" + directory[2];

            return Directory.GetFiles(userFolder + "/Downloads", "*.csv");
        }
        /// <summary>
        /// Filters out file paths with qualifying contents
        /// </summary>
        /// <param name="filePaths"></param>
        /// <returns></returns>
        private List<string> OnlyValidPathsFrom(string[] filePaths)
        {
            List<string> validFiles = new List<string>();

            foreach (string file in filePaths)
            {
                foreach (string line in File.ReadLines(file))
                {
                    if (line.StartsWith("Ic Lot Number"))
                    {
                        validFiles.Add(file);
                        break;
                    }
                }
            }
            return validFiles;
        }
        private List<List<List<string>>> LoadFiles(List<string> filePaths)
        {
            List<List<List<string>>> contents = new List<List<List<string>>>();

            foreach (string file in filePaths)
            {
                contents.Add(new List<List<string>>());

                foreach (string line in File.ReadLines(file))
                {
                    if (line.StartsWith(",") == false)
                        contents[contents.Count - 1].Add(line.Split(new char[] { ',' }).ToList());
                }
            }

            return contents;

        }

    }
}
