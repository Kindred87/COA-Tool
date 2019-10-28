using System;
using System.Collections.Generic;
using System.Text;

namespace CoA_Tool.Console
{
    /// <summary>
    /// Represents a user-navigatable, single-choice menu
    /// </summary>
    class SelectionMenu
    {
        private List<string> MenuOptions;

        private int CurrentSelection = 0;

        public string UserChoice;
        public SelectionMenu(List<string> options, string menuTitle, string centerMessage)
        {
            Util.WriteMessageInCenter(centerMessage);
            MenuOptions = options;
            InitialWrite(menuTitle);
            UserChoice = GetUserChoice();
            RemoveMenu(options.Count);
            Util.RemoveMessageInCenter();
        }
        /// <summary>
        /// Outputs all template options to the console
        /// </summary>
        /// <param name="menuTitle">This string is printed at the top of the menu, indicating what each option represents</param>
        private void InitialWrite(string menuTitle)
        {
            System.Console.SetCursorPosition(0, 30);

            System.Console.WriteLine(menuTitle);

            foreach (string option in MenuOptions)
            {
                System.Console.WriteLine("\t" + option);
            }
        }
        private void RemoveMenu(int numberOfOptions)
        {
            for (int i = 0; i < numberOfOptions + 1; i++)
            {
                System.Console.SetCursorPosition(0, 30 + i);
                System.Console.Write(new string(' ', System.Console.WindowWidth));
            }
        }
        /// <summary>
        /// Handles menu navigation and selection
        /// </summary>
        /// <returns>String value of option</returns>
        private string GetUserChoice()
        {
            string choice;

            do
            {
                HighlightCurrentSelection();
                MenuKeyAction(out choice);
            } while (choice == string.Empty);

            return choice;
        }
        /// <summary>
        /// Rewrites the current selection in green
        /// </summary>
        private void HighlightCurrentSelection()
        {
            System.Console.SetCursorPosition(0, 31 + CurrentSelection);
            System.Console.Write(new string(' ' , System.Console.WindowWidth));

            System.Console.SetCursorPosition(0, 31 + CurrentSelection);
            System.Console.ForegroundColor = ConsoleColor.Green;
            System.Console.Write("\t" + MenuOptions[CurrentSelection]);
            System.Console.ForegroundColor = ConsoleColor.Gray;
        }
        /// <summary>
        /// Rewrites the current selection in gray
        /// </summary>
        private void RemoveHighlightFromCurrentSelection()
        {
            System.Console.SetCursorPosition(0, 31 + CurrentSelection);
            System.Console.Write(new string(' ', System.Console.WindowWidth));

            System.Console.SetCursorPosition(0, 31 + CurrentSelection);
            System.Console.Write("\t" + MenuOptions[CurrentSelection]);
        }
        /// <summary>
        /// Performs a menu-related action based on the key pressed, assinging a value to a string if applicable
        /// </summary>
        private void MenuKeyAction(out string choice)
        {
            choice = string.Empty; // Circumvents assignment error

            switch (System.Console.ReadKey().Key)
            {
                case ConsoleKey.Enter:
                    choice = MenuOptions[CurrentSelection];
                    break;
                case ConsoleKey.RightArrow:
                    choice = MenuOptions[CurrentSelection];
                    break;
                case ConsoleKey.UpArrow:
                    if (CurrentSelection - 1 >= 0)
                    {
                        RemoveHighlightFromCurrentSelection();
                        CurrentSelection--;
                        break;
                    }
                    else
                        break;
                    
                case ConsoleKey.DownArrow:
                    if (CurrentSelection + 1 <= MenuOptions.Count- 1)
                    {
                        RemoveHighlightFromCurrentSelection();
                        CurrentSelection++;
                        break;
                    }
                    else
                        break;
                default:
                    break;
            }
        }
    }
}
