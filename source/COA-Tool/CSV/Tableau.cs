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
        public enum LotDirectory { NewBatch, InProgress, Complete, PreviousBatch, DeletionQueue}

        public List<List<List<string>>> FileContents;

        public string DesktopSubDirectoryPath;

        public Tableau()
        {
            FileContents = new List<List<List<string>>>();
            DesktopSubDirectoryPath = Environment.GetFolderPath(Environment.SpecialFolder.Desktop) + "\\CoAs\\Tableau Files";
        }
        /// <summary>
        /// Retrieves paths of applicable .csv files from the provided directory
        /// </summary>
        /// <returns></returns>  
        private List<string> CSVFilesFrom(string directory)
        {
            return OnlyValidPathsFrom(Directory.GetFiles(directory, "*.csv"));
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
        /// Populates FileContents with data from applicable CSV files
        /// </summary>
        /// <param name="filePaths"></param>
        /// <returns></returns>
        public void Load()
        {
            CleanAllLotDirectories();

            string choice = BatchChoice();

            List<string> filePaths;

            if (choice == "New")
            {
                MoveToDirectory(CSVFilesFrom(Environment.GetFolderPath(Environment.SpecialFolder.UserProfile) + "\\Downloads\\"), LotDirectory.NewBatch);
                filePaths = (CSVFilesFrom(DesktopSubDirectoryPath + "\\New Batch"));
            }
            else if (choice.StartsWith("In")) // The full string contains dynamic values
            {
                filePaths = (CSVFilesFrom(DesktopSubDirectoryPath + "\\In Progress"));
            }
            else //  Load from /Previous Batch
            {
                filePaths = (CSVFilesFrom(DesktopSubDirectoryPath + "\\Previous Batch"));
            }

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
        /// <summary>
        /// Moves given .txt files to one of several pre-determined directories
        /// </summary>
        /// <param name="filePaths"></param>
        /// <param name="targetDirectory"></param>
        public void MoveToDirectory(List<string> filePaths, LotDirectory targetDirectory)
        {
            EnsureDesktopSubDirectoriesExist();

            string desktopDirectoryPath;

            if (targetDirectory == LotDirectory.NewBatch)
                desktopDirectoryPath = "/CoAs/Tableau Files/New Batch/";
            else if(targetDirectory == LotDirectory.InProgress)
                desktopDirectoryPath = "/CoAs/Tableau Files/In Progress/";
            else if (targetDirectory == LotDirectory.Complete)
                desktopDirectoryPath = "/CoAs/Tableau Files/Complete/";
            else if (targetDirectory == LotDirectory.PreviousBatch)
                desktopDirectoryPath = "/CoAs/Tableau Files/Previous Batch/";
            else
                desktopDirectoryPath = "/CoAs/Tableau Files/Deletion Queue/";

            string newFileName;

            foreach(string filePath in filePaths)
            {
                newFileName = File.ReadLines(filePath).ElementAt(1).Split(new char[] { ',' })[3];
                File.Move(filePath, Environment.GetFolderPath(Environment.SpecialFolder.Desktop) + desktopDirectoryPath + newFileName + ".csv" , true);
            }
        }
        /// <summary>
        /// Prompts user to select between a new or previous batch
        /// </summary>
        /// <returns></returns>
        private string BatchChoice()
        {
            List<string> options = new List<string>();

            int amountInProgress = Directory.GetFiles(Environment.GetFolderPath(Environment.SpecialFolder.Desktop) + "/CoAs/Tableau Files/In Progress", "*.csv").Length;

            options.Add("New");
            options.Add("Previous");

            if (amountInProgress > 0)
            {
                options.Insert(1, "In Progess (" + amountInProgress);
                
                if (amountInProgress > 1)
                    options[1] += " files)";
                else
                    options[1] += " file)";
            }
                
           Console.SelectionMenu menu = new Console.SelectionMenu(options, "    Batch:", "Select an option regarding Tableau lot information.");

           return menu.UserChoice;
        }
        /// <summary>
        /// Creates necessary lot information directories if missing
        /// </summary>
        private void EnsureDesktopSubDirectoriesExist()
        {
            if (Directory.Exists(Environment.GetFolderPath(Environment.SpecialFolder.Desktop) + "/CoAs/Tableau Files/New Batch") == false)
                Directory.CreateDirectory(Environment.GetFolderPath(Environment.SpecialFolder.Desktop) + "/CoAs/Tableau Files/New Batch");

            if (Directory.Exists(Environment.GetFolderPath(Environment.SpecialFolder.Desktop) + "/CoAs/Tableau Files/Previous Batch") == false)
                Directory.CreateDirectory(Environment.GetFolderPath(Environment.SpecialFolder.Desktop) + "/CoAs/Tableau Files/Previous Batch");

            if (Directory.Exists(Environment.GetFolderPath(Environment.SpecialFolder.Desktop) + "/CoAs/Tableau Files/In Progress") == false)
                Directory.CreateDirectory(Environment.GetFolderPath(Environment.SpecialFolder.Desktop) + "/CoAs/Tableau Files/In Progress");

            if (Directory.Exists(Environment.GetFolderPath(Environment.SpecialFolder.Desktop) + "/CoAs/Tableau Files/Complete") == false)
                Directory.CreateDirectory(Environment.GetFolderPath(Environment.SpecialFolder.Desktop) + "/CoAs/Tableau Files/Complete");

            if (Directory.Exists(Environment.GetFolderPath(Environment.SpecialFolder.Desktop) + "/CoAs/Tableau Files/Deletion Queue") == false)
                Directory.CreateDirectory(Environment.GetFolderPath(Environment.SpecialFolder.Desktop) + "/CoAs/Tableau Files/Deletion Queue");
        }
        /// <summary>
        /// Searches for and deletes any file with a duplicate sales order
        /// </summary>
        /// <param name="directory"></param>
        private void CleanLotDirectory(LotDirectory directory)
        {

            string fullPath = DesktopSubDirectoryPath;

            if (directory == LotDirectory.NewBatch)
                fullPath += "\\New Batch\\";
            else if (directory == LotDirectory.InProgress)
                fullPath += "\\In Progress\\";
            else if (directory == LotDirectory.PreviousBatch)
                fullPath += "\\Previous Batch\\";
            else if (directory == LotDirectory.Complete)
                fullPath += "\\Complete\\";
            else
                fullPath += "\\Deletion Queue\\";

            HashSet<string> salesOrderSet = new HashSet<string>();
            string salesOrder;

            foreach(string file in CSVFilesFrom(fullPath))
            {
                salesOrder = File.ReadLines(file).ElementAt(1).Split(new char[] { ',' })[3];

                if (salesOrderSet.Contains(salesOrder))
                    File.Delete(file);
                else
                    salesOrderSet.Add(salesOrder);
            }
        }
        /// <summary>
        /// Calls all variations of CleanLotDirectory
        /// </summary>
        private void CleanAllLotDirectories()
        {
            CleanLotDirectory(LotDirectory.NewBatch);
            CleanLotDirectory(LotDirectory.InProgress);
            CleanLotDirectory(LotDirectory.Complete);
            CleanLotDirectory(LotDirectory.PreviousBatch);
            CleanLotDirectory(LotDirectory.DeletionQueue);
        }

    }
}
