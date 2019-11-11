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
                if (DataGroup != DataGroupToSearch.Unassigned && SearchCriteria != string.Empty && SearchColumnOffset >= 0)
                    return true;
                else
                    return false;
            }
        }
        /// <summary>
        /// Indicates that the offset should be from a particular column
        /// </summary>
        public bool OffsetFromSpecificColumn;
        
        /// <summary>
        /// Files or data clusters to search for additional user-requested data
        /// </summary>
        public enum DataGroupToSearch { Unassigned, Titration, Micro };

        private DataGroupToSearch _dataGroup;
        public DataGroupToSearch DataGroup
        {
            get
            {
                return _dataGroup;
            }
            set
            {
                if(SearchColumnOffset > -1) // In the event that SearchColumnOffset is assigned before DataGroup
                {
                    SearchColumnOffset = ++SearchColumnOffset; // Set property is reexecuted to properly assign SearchColumnOffset for the data group
                                                               // value passed to SearchColumnOffset set property is subtracted by one
                }
            }
        }
        private int _searchColumnOffset;
        public int SearchColumnOffset
        {
            get
            {
                return _searchColumnOffset;
            }
            set
            {
                // 0-based in program, 1-based in template file for users' sake
                _searchColumnOffset = value - 1;   

                if(DataGroup == DataGroupToSearch.Micro)
                {
                    
                    if(_searchColumnOffset == 11 || _searchColumnOffset >= 35)
                    {
                        string invalidColumnPrompt = (_searchColumnOffset + 1) + " is not allowed for a target column value for micro data group";
                        List<string> options = new List<string>();
                        options.Add("Continue regardless.");
                        options.Add("Exit application");
                        if (new Console.SelectionMenu(options, "", invalidColumnPrompt).UserChoice == "Exit application")
                            Environment.Exit(0);
                    }
                    else if (_searchColumnOffset > 11)
                    {
                        OffsetFromSpecificColumn = true;
                        _searchColumnOffset -= 13;
                    }
                }
            }
        }

        public string SearchCriteria;

        // Constructor
        public CustomSearch()
        {
            DataGroup = DataGroupToSearch.Unassigned;
            SearchCriteria = string.Empty;
            SearchColumnOffset = 0; // Equates to -1
            OffsetFromSpecificColumn = false;
        }
    }
}
