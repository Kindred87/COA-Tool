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
            List<string> downloadsFilePaths = OnlyValidPathsFrom(GetCSVFilePathsFromDownloads());

            MoveFilesToDirectoryInDesktop(downloadsFilePaths);

            if(BatchChoice() == "New batch")
            {
                // Load from /New Batch
            }
            else //  == "Previous batch"
            {
                // Load from // Previous Batch
            }

            LoadFiles(downloadsFilePaths);
        }
        /// <summary>
        /// Fetches list of csv file paths from the user's download folder
        /// </summary>
        /// <returns></returns>  
        private string[] GetCSVFilePathsFromDownloads()
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
            // A list is used because the number of needed indices is unknown, running a loop to determine that value would invalidate the performance gain
            List<string> validFiles = new List<string>();
            string firstLine;

            foreach (string path in filePaths)
            {
                firstLine = File.ReadLines(path).First();

                if (firstLine.StartsWith("Ic Lot Number"))
                    validFiles.Add(path);
            }
            return validFiles;
        }
        /// <summary>
        /// Reads data from lot information files and stores in a multi-jagged list
        /// </summary>
        /// <param name="filePaths"></param>
        /// <returns></returns>
        public void LoadFiles(List<string> filePaths)
        {
            FileContents = new List<List<List<string>>>();

            foreach (string file in filePaths)
            {
                FileContents.Add(new List<List<string>>());

                foreach (string line in File.ReadLines(file))
                {
                    if (line.StartsWith(",") == false)
                        FileContents[FileContents.Count - 1].Add(line.Split(new char[] { ',' }).ToList());
                }
            }
        }
        private void MoveFilesToDirectoryInDesktop(List<string> filePaths)
        {
            if (Directory.Exists(Environment.GetFolderPath(Environment.SpecialFolder.Desktop) + "/CoAs/Tableau Files/New Batch") == false)
                Directory.CreateDirectory(Environment.GetFolderPath(Environment.SpecialFolder.Desktop) + "/CoAs/Tableau Files/New Batch");

            string newFileName;

            foreach(string filePath in filePaths)
            {
                newFileName = File.ReadAllLines(filePath)[1].Split(new char[] { ',' })[3];
                File.Move(filePath, Environment.GetFolderPath(Environment.SpecialFolder.Desktop) + "/CoAs/Tableau Files/New Batch" + newFileName + ".txt" , true);
            }
        }
        /// <summary>
        /// Prompts user to select between a new or previous batch
        /// </summary>
        /// <returns></returns>
        private string BatchChoice()
        {
            Console.Util.WriteMessageInCenter("Select an option regarding Tableau lot information.");

            string[] options = { "New batch", "Previous batch" };

           Console.SelectionMenu menu = new Console.SelectionMenu(options, "    Use:");

            Console.Util.RemoveMessageInCenter();

            return menu.UserChoice;
        }

    }
}
