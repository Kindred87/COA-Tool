using System;
using System.Collections.Generic;
using System.Text;
using System.IO;
using System.Linq;

namespace CoA_Tool.CSV
{
    /// <summary>
    /// Handles location and loading of downloaded Tableau data
    /// </summary>
    class Tableau
    {
        public List<List<List<string>>> FileContents;
        public Tableau()
        {
            List<string> filePaths = OnlyValidPathsFrom(GetCSVFilePaths());

            FileContents = LoadFiles(filePaths);
            MoveFilesToDesktop(filePaths);
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
        /// <summary>
        /// Reads data from lot information files and stores in a multi-jagged list
        /// </summary>
        /// <param name="filePaths"></param>
        /// <returns></returns>
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
        private void MoveFilesToDesktop(List<string> filePaths)
        {
            if (Directory.Exists(Environment.GetFolderPath(Environment.SpecialFolder.Desktop) + "/CoAs/Tableau Files") == false)
                Directory.CreateDirectory(Environment.GetFolderPath(Environment.SpecialFolder.Desktop) + "/CoAs/Tableau Files");

            string newFileName;

            foreach(string filePath in filePaths)
            {
                newFileName = File.ReadAllLines(filePath)[1].Split(new char[] { ',' })[3];
                File.Move(filePath, Environment.GetFolderPath(Environment.SpecialFolder.Desktop) + "/CoAs/Tableau Files/" + newFileName + ".txt" , true);
            }
        }

    }
}
