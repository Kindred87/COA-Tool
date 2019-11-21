using System;
using System.Collections.Generic;
using System.Text;
using CoA_Tool.Utility;

namespace CoA_Tool.ConsoleInteraction
{
    /// <summary>
    /// Represents a user-navigable, single-choice menu
    /// </summary>
    class SelectionMenu
    {
        // Lists
        /// <summary>
        /// Contains menu items arranged as a grid in terms of [x][y]
        /// </summary>
        private List<List<string>> MenuGrid;

        // Arrays
        /// <summary>
        /// The zero-based location of the highlighted item within the menu grid in terms of [x][y] 
        /// </summary>
        private int[] SelectionCoordinates;

        // Integers
        /// <summary>
        /// Used to determine the maximum number of menu columns given the console window's width
        /// </summary>
        private int ColumnLimitForWindow
        {
            get
            {
                return System.Console.WindowWidth / GridColumnWidth ;
            }
        }
        /// <summary>
        /// The amount of horizontal whitespace alloted for each menu item
        /// </summary>
        private int GridColumnWidth;
        /// <summary>
        /// Stores the console window's width for resizing detection.
        /// </summary>
        private int WindowWidthOnLastUpdate;
        /// <summary>
        /// Stores the console window's width for resizing detection.
        /// </summary>
        private int WindowHeightOnLastUpdate;
        private int TopRowOfGrid
        {
            get
            {
                return (int)(Console.WindowHeight * 0.5) + 4;
            }
        }
        /// <summary>
        /// The maximum height of each column within the menu grid in terms of indices.
        /// </summary>
        private int ItemsPerColumn = 10;
        /// <summary>
        /// The amount of whitespace to be added to the GridColumnWidth field
        /// </summary>
        private int SpaceBetweenColumns = 2;

        // Strings
        /// <summary>
        /// The full string of the menu option selected by the user
        /// </summary>
        public string UserChoice;

        /// <summary>
        /// This value sits at the upper-left corner of the menu grid
        /// </summary>
        private string ChoiceHeader;
        /// <summary>
        /// The message to be displayed in the center of the console window
        /// </summary>
        private string CenterMessage;

        // Constructor
        // Operates as a sequence for the menu's lifespan to simplify choice retrieval in other parts of the program.
        public SelectionMenu(List<string> options, string optionsHeader, string centerMessage)
        {
            ChoiceHeader = optionsHeader;
            CenterMessage = centerMessage;
            MenuOptions = options;

            MenuGrid = OptionsToMenuGrid(options);

            GridColumnWidth = SetColumnWidth(MenuGrid);

            SelectionCoordinates = new int[] { 0, 0 };

            SaveWindowDimensions();

            UpdateMenu();

            UserChoice = MenuItemDesiredByUser();

            RemoveMenuContent();
        }

        // Public methods

        // Private methods
        /// <summary>
        /// Outputs all template options to the console
        /// </summary>
        /// <param name="menuTitle">This string is printed at the top of the menu, indicating what each option represents</param>
        private void UpdateMenu()
        {
            RemoveMenuContent();

            // Quick solution to prevent out-of-range exception following console window re-sizing
            SelectionCoordinates[0] = 0; 
            SelectionCoordinates[1] = 0; 

            ConsoleOps.WriteMessageInCenter(CenterMessage);

            Console.SetCursorPosition(0, TopRowOfGrid);

            Console.Write(ChoiceHeader);

            MenuGridToConsole();
        }
        /// <summary>
        /// Handles menu navigation and selection
        /// </summary>
        /// <returns>String value of option</returns>
        private string MenuItemDesiredByUser()
        {
            string chosenMenuItem;
            bool userSubmittedChoice;

            do
            {
                UpdateMenuIfWindowResized();
                
                HighlightCurrentSelection();

                MenuKeyAction(out userSubmittedChoice, out chosenMenuItem);

            } while (userSubmittedChoice == false);

            return chosenMenuItem;
        }
        /// <summary>
        /// Rewrites the current selection in green
        /// </summary>
        private void HighlightCurrentSelection()
        {
            UpdateMenuIfWindowResized();

            Console.SetCursorPosition(SelectionCoordinates[0] * GridColumnWidth, TopRowOfGrid + 1 + SelectionCoordinates[1]);
            Console.ForegroundColor = ConsoleColor.Green;
            Console.Write("\t" + MenuGrid[SelectionCoordinates[0]][SelectionCoordinates[1]]);
            Console.ForegroundColor = ConsoleColor.Gray;
        }
        /// <summary>
        /// Rewrites the current selection in gray
        /// </summary>
        private void RemoveHighlightFromCurrentSelection()
        {
            UpdateMenuIfWindowResized();

            Console.SetCursorPosition(SelectionCoordinates[0] * GridColumnWidth, TopRowOfGrid + 1 + SelectionCoordinates[1]);
            Console.Write("\t" + MenuGrid[SelectionCoordinates[0]][SelectionCoordinates[1]]);
        }
        /// <summary>
        /// <para>Performs a menu-related action based on the key pressed.</para>
        /// Out parameters indicate if the menu action was a choice submission and the string value of that choice.
        /// </summary>
        /// <param name="itemChoiceSubmitted">Indicates if the menu action was a choice submission.</param>
        /// <param name="itemChoice">If the menu action was a choice submission, this is the value of that choice.</param>
        private void MenuKeyAction(out bool itemChoiceSubmitted, out string itemChoice)
        {
            switch (Console.ReadKey(true).Key)
            {
                case ConsoleKey.Enter:
                    itemChoiceSubmitted = true;
                    itemChoice = MenuOptions[SelectionCoordinates[0] * 10 + SelectionCoordinates[1]];
                    break;
                case ConsoleKey.Tab:
                    itemChoiceSubmitted = true;
                    itemChoice = MenuOptions[SelectionCoordinates[0] * 10 + SelectionCoordinates[1]];
                    break;
                case ConsoleKey.RightArrow:
                    MoveRight();
                    itemChoiceSubmitted = false;
                    itemChoice = string.Empty;
                    break;
                case ConsoleKey.D:
                    MoveRight();
                    itemChoiceSubmitted = false;
                    itemChoice = string.Empty;
                    break;
                case ConsoleKey.LeftArrow:
                    MoveLeft();
                    itemChoiceSubmitted = false;
                    itemChoice = string.Empty;
                    break;
                case ConsoleKey.A:
                    MoveLeft();
                    itemChoiceSubmitted = false;
                    itemChoice = string.Empty;
                    break;
                case ConsoleKey.UpArrow:
                    MoveUp();
                    itemChoiceSubmitted = false;
                    itemChoice = string.Empty;
                    break;
                case ConsoleKey.W:
                    MoveUp();
                    itemChoiceSubmitted = false;
                    itemChoice = string.Empty;
                    break;
                case ConsoleKey.DownArrow:
                    MoveDown();
                    itemChoiceSubmitted = false;
                    itemChoice = string.Empty;
                    break;
                case ConsoleKey.S:
                    MoveDown();
                    itemChoiceSubmitted = false;
                    itemChoice = string.Empty;
                    break;
                default:
                    itemChoiceSubmitted = false;
                    itemChoice = string.Empty;
                    break;
            }
        }
        /// <summary>
        /// Modifies CurrentSelection[0] to highlight the menu option to the right of the current selection.
        /// Includes wrapping logic.
        /// </summary>
        private void MoveRight()
        {
            // If there are no columns to the right
            if(ColumnLimitForWindow <= SelectionCoordinates[0] + 1 || MenuGrid.Count <= SelectionCoordinates[0] + 1)
            {
                RemoveHighlightFromCurrentSelection();
                SelectionCoordinates[0] = 0;
            }
            // If there is no value on the same row to the immediate right
            else if(MenuGrid[SelectionCoordinates[0] + 1].Count - 1 < SelectionCoordinates[1])
            {
                RemoveHighlightFromCurrentSelection();
                SelectionCoordinates[0]++;
                SelectionCoordinates[1] = MenuGrid[SelectionCoordinates[0]].Count - 1;
            }
            else
            {
                RemoveHighlightFromCurrentSelection();
                SelectionCoordinates[0]++;
            }
        }
        /// <summary>
        /// Modifies CurrentSelection[0] to highlight the menu option to the left of the current selection.
        /// Includes wrapping logic.
        /// </summary>
        private void MoveLeft()
        {
            // If there are no columns to the left
            if (SelectionCoordinates[0] - 1 < 0)
            {
                RemoveHighlightFromCurrentSelection();

                if(MenuGrid.Count > ColumnLimitForWindow)
                {
                    SelectionCoordinates[0] = ColumnLimitForWindow - 1;

                }
                else
                {
                    SelectionCoordinates[0] = MenuGrid.Count - 1;
                }

                // If there are no values on the same row in the last column of the grid
                if(MenuGrid[SelectionCoordinates[0]].Count - 1 < SelectionCoordinates[1])
                {
                    SelectionCoordinates[1] = MenuGrid[SelectionCoordinates[0]].Count - 1;
                }
            }
            else
            {
                RemoveHighlightFromCurrentSelection();
                SelectionCoordinates[0]--;
            }
        }
        /// <summary>
        /// Modifies CurrentSelection[0] to highlight the menu option above the current selection.
        /// Includes wrapping logic.
        /// </summary>
        private void MoveUp()
        {
            // If there is no row above
            if (SelectionCoordinates[1] - 1 >= 0)
            {
                RemoveHighlightFromCurrentSelection();
                SelectionCoordinates[1]--;
            }
            else
            {
                RemoveHighlightFromCurrentSelection();
                SelectionCoordinates[1] = MenuGrid[SelectionCoordinates[0]].Count - 1;
            }
        }
        /// <summary>
        /// Modifies CurrentSelection[0] to highlight the menu option below the current selection.
        /// Includes wrapping logic.
        /// </summary>
        private void MoveDown()
        {
            // If a row is below
            if (SelectionCoordinates[1] + 1 <= MenuGrid[SelectionCoordinates[0]].Count - 1)
            {
                RemoveHighlightFromCurrentSelection();
                SelectionCoordinates[1]++;
            }
            else
            {
                RemoveHighlightFromCurrentSelection();
                SelectionCoordinates[1] = 0;
            }
        }
        /// <summary>
        /// Indicates whether the console window dimensions has changed since previously checked
        /// </summary>
        /// <returns></returns>
        private bool WindowSizeChanged()
        {
            if (Console.WindowWidth != WindowWidthOnLastUpdate || System.Console.WindowHeight != WindowHeightOnLastUpdate)
            {
                return true;
            }
            else
            {
                return false;
            }
        }
        /// <summary>
        /// Sets window dimension fields to inform window resizing detection
        /// </summary>
        private void SaveWindowDimensions()
        {
            WindowWidthOnLastUpdate = Console.WindowWidth;
            WindowHeightOnLastUpdate = Console.WindowHeight;
        }
        /// <summary>
        /// Determines column width based on the length of the longest menu item and the 
        /// SpaceBetweenColumns variable.
        /// </summary>
        private int SetColumnWidth (List<List<string>> menuGrid)
        {
            int lengthOfLongestString = 0;

            foreach (List<string> column in MenuGrid)
            {
                foreach(string item in column)
                {
                    if(item.Length > lengthOfLongestString)
                    {
                        lengthOfLongestString = item.Length;
                    }
                }
            }

            return lengthOfLongestString + SpaceBetweenColumns;
        }
        /// <summary>
        /// Distributes a linear list of options to a jagged list representing the menu grid.
        /// </summary>
        /// <param name="optionsAsLinear">The linearly-arranged options to distribute.</param>
        /// <returns></returns>
        private List<List<string>> OptionsToMenuGrid(List<string> optionsAsLinear)
        {
            int gridColumnQuantity = NumberOfColumnsInGrid(optionsAsLinear.Count);

            List<List<string>> optionsAsGrid = new List<List<string>>();

            for(int columnIterator = 0; columnIterator < gridColumnQuantity; columnIterator++)
            {
                int rowsWithinColumn;

                // Determine if the current column contains less than the maximum allowed
                if(optionsAsLinear.Count - 1 - (columnIterator * ItemsPerColumn) < ItemsPerColumn)
                {
                    rowsWithinColumn = optionsAsLinear.Count - (columnIterator * ItemsPerColumn);
                }
                else
                {
                    rowsWithinColumn = ItemsPerColumn;
                }

                // Prepare next column for distribution
                optionsAsGrid.Add(new List<string>());
                int columnIndex = optionsAsGrid.Count - 1;

                for (int rowIterator = 0; rowIterator < rowsWithinColumn; rowIterator++)
                {
                    optionsAsGrid[columnIndex].Add(optionsAsLinear[columnIterator * 10 + rowIterator]);
                }
            }

            return optionsAsGrid;
        }
        /// <summary>
        /// Determines how many columns the menu grid should contain.
        /// Based from ItemsPerColumn.
        /// </summary>
        /// <param name="numberOfItemsInGrid"></param>
        /// <returns></returns>
        private int NumberOfColumnsInGrid(int numberOfItemsInGrid)
        {
            if(numberOfItemsInGrid % ItemsPerColumn == 0)
            {
                return numberOfItemsInGrid / ItemsPerColumn;
            }
            else
            {
                return numberOfItemsInGrid / ItemsPerColumn + 1;
            }
        }
        /// <summary>
        /// Removes the menu's content from the console window
        /// </summary>
        private void RemoveMenuContent()
        {
            ConsoleOps.RemoveMessageInCenter();

            for(int i = TopRowOfGrid; i < ItemsPerColumn; i++)
            {
                Console.Write(new string(' ', Console.WindowWidth));
            }
            
            Console.CursorVisible = true;
            Console.SetCursorPosition(0, 0);
            Console.CursorVisible = false;
        }
        private void MenuGridToConsole()
        {
            int columnCount;
            if(MenuGrid.Count <= ColumnLimitForWindow)
            {
                columnCount = MenuGrid.Count;
            }
            else
            {
                columnCount = ColumnLimitForWindow;
            }
            for(int columnIterator = 0; columnIterator < columnCount; columnIterator++)
            {
                for(int rowIterator = 0; rowIterator < MenuGrid[columnIterator].Count; rowIterator++)
                {
                    Console.SetCursorPosition(columnIterator * GridColumnWidth, TopRowOfGrid + 1 + rowIterator);
                    Console.Write("\t" + MenuGrid[columnIterator][rowIterator]);
                }
            }
        }
        private void UpdateMenuIfWindowResized()
        {
            if(WindowSizeChanged())
            {
                UpdateMenu();
                SaveWindowDimensions();
            }
        }
    }
}
