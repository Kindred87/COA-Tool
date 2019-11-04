using System;
using System.Collections.Generic;
using System.Text;

namespace CoA_Tool.Templates
{
    /// <summary>
    /// Contains information for a user-specified search to be performed in tandem with an algorithm
    /// </summary>
    class CustomSearch
    {
        /// <summary>
        /// Is true when all necessary variables have been set
        /// </summary>
        public bool ValidSearch // Used in lieu of a constructor due to parsing being line-by-line and data being in no set order
        {
            get
            {
                if (DataGroup != DataGroupToSearch.Unassigned && SearchCriteria != string.Empty && TargetColumn >= 0)
                    return true;
                else
                    return false;
            }
        }
        /// <summary>
        /// Files or data clusters to search for additional user-requested data
        /// </summary>
        public enum DataGroupToSearch { Unassigned, Titration, Micro };

        public DataGroupToSearch DataGroup;

        public int TargetColumn
        {
            get
            {
                return TargetColumn;
            }
            set
            {
                // 0-based in program, 1-based in template file for users' sake
                TargetColumn = value - 1;   
            }
        }

        public string SearchCriteria;

        // Constructor
        public CustomSearch()
        {
            DataGroup = DataGroupToSearch.Unassigned;
            SearchCriteria = string.Empty;
            TargetColumn = -1;
        }
        /// <summary>
        /// Sets DataGroup from a valid string, offers user to exit if string is invalid
        /// </summary>
        /// <param name="fromTemplateFile"></param>
        public void SetDataGroup(string fromTemplateFile)
        {
            if(fromTemplateFile.ToLower() == "micro")
            {

            }
            else if(fromTemplateFile.ToLower() == "dressing")
            {

            }
            else
            {
                string prompt = "Selected template specifies " + fromTemplateFile + " for \"Data-Group To Search\", this is invalid and your search criteria will not be used";
                
                List<string> options = new List<string>();
                options.Add("Continue regardless");
                
                options.Add("Exit program");
                Console.SelectionMenu menu = new Console.SelectionMenu(options, "Choose to", prompt);

                if(menu.UserChoice == "Exit program")
                {
                    Environment.Exit(0);
                }
                
            }
        }
    }
}
