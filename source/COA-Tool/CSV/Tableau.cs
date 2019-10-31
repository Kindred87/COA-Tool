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
        // Class variables
        /// <summary>
        /// Associated with sub-directories within "Desktop\\CoAs\\Tableau Files"
        /// </summary>
        public enum LotDirectory { NewBatch, InProgress, Complete, PreviousBatch, DeletionQueue}
        /// <summary>
        /// Data loaded from Tableau files
        /// </summary>
        public List<List<List<string>>> FileContents;
        /// <summary>
        /// "Desktop\\CoAs\\Tableau Files", assigned in constructor
        /// </summary>
        public string DesktopSubDirectoryPath;

        // Constructor
        public Tableau()
        {
            FileContents = new List<List<List<string>>>();
            EnsureDesktopSubDirectoriesExist();
            DesktopSubDirectoryPath = Environment.GetFolderPath(Environment.SpecialFolder.Desktop) + "\\CoAs\\Tableau Files";
        }
        // Public methods
        /// <summary>
        /// Populates FileContents with data from applicable CSV files
        /// </summary>
        /// <returns></returns>
        public void Load()
        {
            CleanAllLotDirectories();

            string choice = BatchChoice();

            List<string> filePaths;

            if (choice == "New")
            {
                MoveToDirectory(CSVFilesFrom(Environment.GetFolderPath(Environment.SpecialFolder.UserProfile) + "\\Downloads\\"), LotDirectory.NewBatch);
                filePaths = (CSVFilesFrom(DesktopSubDirectoryPath + "\\1) New Batch"));
            }
            else if (choice.StartsWith("In")) // The full string contains dynamic values, "In Progress (n file/s)"
            {
                filePaths = (CSVFilesFrom(DesktopSubDirectoryPath + "\\2) Current Batch\\1) In Progress"));
            }
            else //  Load from \\3) Previous Batch
            {
                filePaths = (CSVFilesFrom(DesktopSubDirectoryPath + "\\3) Previous Batch"));
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
        /// Moves given files to one of several pre-determined directories
        /// </summary>
        /// <param name="filePaths">Each entry represents a file's absolute path</param>
        /// <param name="targetDirectory">The output directory</param>
        public void MoveToDirectory(List<string> filePaths, LotDirectory targetDirectory)
        {
            EnsureDesktopSubDirectoriesExist();

            string nextSubDirectory;

            if (targetDirectory == LotDirectory.NewBatch)
                nextSubDirectory = "\\1) New Batch\\";
            else if(targetDirectory == LotDirectory.InProgress)
                nextSubDirectory = "\\2) Current Batch\\1) In Progress\\";
            else if (targetDirectory == LotDirectory.Complete)
                nextSubDirectory = "\\3) Complete\\";
            else if (targetDirectory == LotDirectory.PreviousBatch)
                nextSubDirectory = "\\2) CUrrent Batch\\2) Previous Batch\\";
            else
                nextSubDirectory = "\\4) Deletion Queue\\";

            string newFileName;

            foreach(string filePath in filePaths)
            {
                newFileName = File.ReadLines(filePath).ElementAt(1).Split(new char[] { ',' })[3];
                File.Move(filePath, DesktopSubDirectoryPath + nextSubDirectory + newFileName + ".csv" , true);
            }
        }

        // Private methods
        /// <summary>
        /// Prompts user to select which batch of tableau files to generate CoAs from
        /// </summary>
        /// <returns></returns>
        private string BatchChoice()
        {
            List<string> options = new List<string>();

            int amountInProgress = Directory.GetFiles(DesktopSubDirectoryPath + "\\2) Current Batch\\1) In Progress", "*.csv").Length;

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
            if (Directory.Exists(DesktopSubDirectoryPath + "\\1) New Batch") == false)
                Directory.CreateDirectory(DesktopSubDirectoryPath + "\\1) New Batch");

            if (Directory.Exists(DesktopSubDirectoryPath + "\\3) Previous Batch") == false)
                Directory.CreateDirectory(DesktopSubDirectoryPath + "\\3) Previous Batch");

            if (Directory.Exists(DesktopSubDirectoryPath + "\\2) Current Batch\\1) In Progress") == false)
                Directory.CreateDirectory(DesktopSubDirectoryPath + "\\2) Current Batch\\1) In Progress");

            if (Directory.Exists(DesktopSubDirectoryPath + "\\2) Current Batch\\2) Complete") == false)
                Directory.CreateDirectory(DesktopSubDirectoryPath + "\\2) Current Batch\\2) Complete");

            if (Directory.Exists(DesktopSubDirectoryPath + "\\4) Deletion Queue") == false)
                Directory.CreateDirectory(DesktopSubDirectoryPath + "\\4) Deletion Queue");
        }
        /// <summary>
        /// Searches for and deletes any file with a duplicate sales order
        /// </summary>
        /// <param name="directory">Directory to clean</param>
        private void CleanLotDirectory(LotDirectory directory)
        {

            string fullPath = DesktopSubDirectoryPath;

            if (directory == LotDirectory.NewBatch)
                fullPath += "\\1) New Batch\\";
            else if (directory == LotDirectory.InProgress)
                fullPath += "\\2) Current Batch\\1) In Progress\\";
            else if (directory == LotDirectory.Complete)
                fullPath += "\\2) Current Batch\\2) Complete\\";
            else if (directory == LotDirectory.PreviousBatch)
                fullPath += "\\3) Previous Batch\\";
            else
                fullPath += "\\4) Deletion Queue\\";

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
        /// Calls all variations of CleanLotDirectory (Searches for and deletes any file with a duplicate sales order)
        /// </summary>
        private void CleanAllLotDirectories()
        {
            CleanLotDirectory(LotDirectory.NewBatch);
            CleanLotDirectory(LotDirectory.InProgress);
            CleanLotDirectory(LotDirectory.Complete);
            CleanLotDirectory(LotDirectory.PreviousBatch);
            CleanLotDirectory(LotDirectory.DeletionQueue);
        }
        /// <summary>
        /// Retrieves paths of applicable .csv files from the provided directory.  Does not search sub-directories.
        /// </summary>
        /// <param name="directory">Directory to search within</param>
        /// <returns></returns>  
        private List<string> CSVFilesFrom(string directory)
        {
            return OnlyValidPathsFrom(Directory.GetFiles(directory, "*.csv"));
        }
        /// <summary>
        /// Filters out file paths without matching criteria
        /// </summary>
        /// <param name="filePaths">Array of each file's path</param>
        /// <returns></returns>
        private List<string> OnlyValidPathsFrom(string[] filePaths)
        {
            // A list is used because the number of needed indices is unknown
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

    }
}
