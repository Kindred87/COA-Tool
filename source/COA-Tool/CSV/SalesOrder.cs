using System;
using System.Collections.Generic;
using System.Text;
using System.IO;

namespace CoA_Tool.CSV
{
    class SalesOrder
    {
        /// <summary>
        /// The sales order number
        /// </summary>
        public string OrderNumber;

        /// <summary>
        /// Contains all lots for the sales order
        /// </summary>
        public List<string> Lots;

        /// <summary>
        /// Demonstrates if lots and order numbers have been assigned
        /// </summary>
        public bool ValidSalesOrder
        {
            get
            {
                if(OrderNumber != string.Empty && Lots.Count > 0)
                {
                    return true;
                }
                else
                {
                    return false;
                }
            }
        }

        public SalesOrder(string targetFilePath)
        {
            Lots = new List<string>();
            OrderNumber = string.Empty;
            LoadData(targetFilePath);
        }
        /// <summary>
        /// Parses lot information file and sets Lots and OrderNumber
        /// </summary>
        /// <param name="filePath"></param>
        private void LoadData(string filePath)
        {
            try
            {
                foreach(string line in File.ReadLines(filePath))
                {
                    if(line.StartsWith(",") == false && line.StartsWith("Ic") == false)
                    {
                        string[] delimitedLine = line.Split(new char[] { ',' });

                        if (delimitedLine.Length == 6)
                        {
                            Lots.Add(delimitedLine[0]);
                            OrderNumber = delimitedLine[3];
                        }
                    }
                    else
                    {
                        continue;
                    }
                }
            }
            catch(IOException)
            {
                string[] delimitedPath = filePath.Split(new char[] { '/', '\\' });

                string fileName = delimitedPath[delimitedPath.Length - 1];

                List<string> menuOptions = new List<string>();
                menuOptions.Add("Skip file");
                menuOptions.Add("Reload file");

                if(new Console.SelectionMenu(menuOptions, "", fileName + " is being accessed by another program.  Skip this file or try loading it again?").UserChoice == "Reload file")
                {
                    LoadData(filePath);
                }
                else
                {
                    menuOptions.Clear();
                    menuOptions.Add("Yes");
                    menuOptions.Add("No");

                    if (new Console.SelectionMenu(menuOptions, "", "Are you sure you want to skip loading " + fileName + "?").UserChoice == "No")
                    {
                        LoadData(filePath);
                    }
                    else
                    {
                        return;
                    }
                    
                }
            }
        }
    }
}
