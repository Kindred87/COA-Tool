using System;
using System.Collections.Generic;
using System.Text;
using System.IO;
using System.Linq;

namespace CoA_Tool.CSV
{
    /// <summary>
    /// Handles location and loading of downloaded TableauData data
    /// </summary>
    class TableauData
    {
        // Enums
        /// <summary>
        /// Associated with sub-directories within "Desktop\\CoAs\\TableauData Files"
        /// </summary>
        public enum LotDirectory { NewBatch, InProgress, Complete, PreviousBatch, DeletionQueue}

        // Lists
        /// <summary>
        /// Each index represents a unique sales order
        /// </summary>
        public List<SalesOrder> SalesOrders;

        // String
        /// <summary>
        /// "Desktop\\CoAs\\Tableau Files", assigned in constructor
        /// </summary>
        public string DesktopSubDirectoryPath;
        public string NewBatchPath;
        public string InProgressCurrentBatchPath;
        public string CurrentCompleteBatchPath;
        public string PreviousBatchPath;
        public string DeletionQueuePath;

        // Constructor
        public TableauData()
        {
            SalesOrders = new List<SalesOrder>();
            DesktopSubDirectoryPath = Environment.GetFolderPath(Environment.SpecialFolder.Desktop) + "\\CoAs\\Tableau Files";
            NewBatchPath = DesktopSubDirectoryPath + "\\1) New Batch\\";
            InProgressCurrentBatchPath = DesktopSubDirectoryPath + "\\2) Current Batch\\1) In Progress\\";
            CurrentCompleteBatchPath = DesktopSubDirectoryPath + "\\2) Current Batch\\2) Complete\\";
            PreviousBatchPath = DesktopSubDirectoryPath + "\\3) Previous Batch\\";
            DeletionQueuePath = DesktopSubDirectoryPath + "\\4) Deletion Queue\\";
            EnsureDesktopSubDirectoriesExist();
        }

        // Public methods
        /// <summary>
        /// Populates SalesOrders with sale order
        /// </summary>
        /// <returns></returns>
        public void Load()
        {
            CleanAllLotDirectories();

            string choice = BatchChoice();

            List<string> filePaths;

            if (choice == "New")
            {
                MoveBetweenDirectoriesMulti(CSVFilesFrom(Environment.GetFolderPath(Environment.SpecialFolder.UserProfile) + "\\Downloads\\"), LotDirectory.NewBatch);
                filePaths = (CSVFilesFrom(NewBatchPath));
            }
            else if (choice.StartsWith("In")) // The full string contains dynamic values, "In Progress (n file/s)"
            {
                filePaths = (CSVFilesFrom(InProgressCurrentBatchPath));
            }
            else //  Load from \\3) Previous Batch
            {
                filePaths = CSVFilesFrom(PreviousBatchPath);
                MoveBetweenDirectoriesMulti(filePaths, LotDirectory.NewBatch);
                filePaths = CSVFilesFrom(NewBatchPath);
            }

            foreach (string file in filePaths)
            {
                SalesOrders.Add(new SalesOrder(file));
            }
        }
        /// <summary>
        /// Moves given files to one of several pre-determined directories
        /// </summary>
        /// <param name="filePaths">Each entry represents a file's absolute path</param>
        /// <param name="targetDirectory">The output directory</param>
        public void MoveBetweenDirectoriesMulti(List<string> filePaths, LotDirectory targetDirectory)
        {
            foreach(string filePath in filePaths)
            {
                MoveBetweenDirectoriesSingle(filePath, targetDirectory);
            }
        }
        /// <summary>
        /// Moves a given file to one of several pre-determined directories
        /// </summary>
        /// <param name="filePaths">Each entry represents a file's absolute path</param>
        /// <param name="targetDirectory">The output directory</param>
        public void MoveBetweenDirectoriesSingle(string filePath, LotDirectory targetDirectory)
        {
            EnsureDesktopSubDirectoriesExist();

            string directory;

            if (targetDirectory == LotDirectory.NewBatch)
            {
                directory = NewBatchPath;
            }
            else if (targetDirectory == LotDirectory.InProgress)
            {
                directory = InProgressCurrentBatchPath;
            }
            else if (targetDirectory == LotDirectory.Complete)
            {
                directory = CurrentCompleteBatchPath;
            }
            else if (targetDirectory == LotDirectory.PreviousBatch)
            {
                directory = PreviousBatchPath;
            }
            else
            {
                directory = DeletionQueuePath;
            }

            string newFileName;
            
            newFileName = File.ReadLines(filePath).ElementAt(1).Split(new char[] { ',' })[3];
            File.Move(filePath, directory + newFileName + ".csv", true);
            
        }
        /// <summary>
        /// Removes sales orders with missing information from SalesOrders
        /// </summary>
        public void RemoveInvalidSalesOrders()
        {
            for(int i = 0; i < SalesOrders.Count - 1; i++)
            {
                if(SalesOrders[i].ValidSalesOrder == false)
                {
                    SalesOrders.RemoveAt(i);
                }
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
                {
                    options[1] += " files)";
                }
                else
                {
                    options[1] += " file)";
                }
            }
                
           ConsoleInteraction.SelectionMenu menu = new ConsoleInteraction.SelectionMenu(options, "    Batch:", "Select an option regarding TableauData lot information.");

           return menu.UserChoice;
        }
        /// <summary>
        /// Creates necessary lot information directories if missing
        /// </summary>
        private void EnsureDesktopSubDirectoriesExist()
        {
            if (Directory.Exists(DesktopSubDirectoryPath) == false)
                Directory.CreateDirectory(DesktopSubDirectoryPath);

            if (Directory.Exists(NewBatchPath) == false)
                Directory.CreateDirectory(NewBatchPath);

            if (Directory.Exists(PreviousBatchPath) == false)
                Directory.CreateDirectory(PreviousBatchPath);

            if (Directory.Exists(InProgressCurrentBatchPath) == false)
                Directory.CreateDirectory(InProgressCurrentBatchPath);

            if (Directory.Exists(CurrentCompleteBatchPath) == false)
                Directory.CreateDirectory(CurrentCompleteBatchPath);

            if (Directory.Exists(DeletionQueuePath) == false)
                Directory.CreateDirectory(DeletionQueuePath);
        }
        /// <summary>
        /// Searches for and deletes any file with a duplicate sales order
        /// </summary>
        /// <param name="directory">Directory to clean</param>
        private void CleanLotDirectory(LotDirectory directory)
        {
            string targetPath = DesktopSubDirectoryPath;

            if (directory == LotDirectory.NewBatch)
            {
                targetPath = NewBatchPath;
            }
            else if (directory == LotDirectory.InProgress)
            {
                targetPath = InProgressCurrentBatchPath;
            }
            else if (directory == LotDirectory.Complete)
            {
                targetPath = CurrentCompleteBatchPath;
            }
            else if (directory == LotDirectory.PreviousBatch)
            {
                targetPath = PreviousBatchPath;
            }
            else
            {
                targetPath = DeletionQueuePath;
            }

            HashSet<string> salesOrderSet = new HashSet<string>();
            string salesOrder = string.Empty;

            foreach(string file in CSVFilesFrom(targetPath))
            {
                string secondLine = File.ReadLines(file).ElementAt(1);
                if (secondLine != string.Empty)
                    salesOrder = secondLine.Split(new char[] { ',' })[3];
                else
                {
                    
                    foreach(string line in File.ReadLines(file))
                    {
                        if (line.Contains("Qty"))
                            salesOrder = line.Split(new char[] { ',' })[3];
                    }
                }

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
        public List<string> CSVFilesFrom(string directory)
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

            foreach (string path in filePaths)
            {
                if(IsFileValid(path))
                {
                    validFiles.Add(path);
                }
            }
            return validFiles;
        }
        /// <summary>
        /// Checks if file contains qualifying criteria, includes IOException handling
        /// </summary>
        /// <param name="filePath"></param>
        /// <returns></returns>
        private bool IsFileValid(string filePath)
        {
            string firstLine;

            try
            {
                firstLine = File.ReadLines(filePath).First();

                if (firstLine.StartsWith("Ic Lot Number"))
                {
                    return true;
                }
                else
                {
                    return false;
                }
            }
            catch (IOException)
            {
                string[] delimitedPath = filePath.Split(new char[] { '/', '\\' });

                string fileName = delimitedPath[delimitedPath.Length - 1];

                List<string> menuOptions = new List<string>();
                menuOptions.Add("Skip file");
                menuOptions.Add("Reload file");

                if (new ConsoleInteraction.SelectionMenu(menuOptions, "", fileName + " is being accessed by another program.  Skip this file or try loading it again?").UserChoice == "Reload file")
                {
                    return IsFileValid(filePath);
                }
                else
                {
                    menuOptions.Clear();
                    menuOptions.Add("Yes");
                    menuOptions.Add("No");

                    if(new ConsoleInteraction.SelectionMenu(menuOptions, "", "Are you sure you want to skip loading " + fileName + "?").UserChoice == "No")
                    {
                        return IsFileValid(filePath);
                    }
                    else
                    {
                        return false;
                    }
                    
                }
            }
        }

    }
}
