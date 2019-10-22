using System;
using System.Collections.Generic;
using System.Text;

namespace CoA_Tool.Templates
{
    class SelectionMenu
    {
        private string[] TemplateOptions;

        private int CurrentSelection = 0;

        public string UserChoice;
        public SelectionMenu(string[] options)
        {
            TemplateOptions = options;
            InitialWrite();
            UserChoice = GetUserChoice();
        }
        /// <summary>
        /// Outputs all template options to the console
        /// </summary>
        private void InitialWrite()
        {
            System.Console.SetCursorPosition(0, 30);

            System.Console.WriteLine("Templates: ");

            foreach (string option in TemplateOptions)
            {
                System.Console.WriteLine("\t" + option);
            }
        }
        /// <summary>
        /// Handles menu navigation and selection, returns name of chosen template
        /// </summary>
        /// <returns></returns>
        private string GetUserChoice()
        {
            string choice = string.Empty;

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
            System.Console.Write("\t" + TemplateOptions[CurrentSelection]);
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
            System.Console.Write("\t" + TemplateOptions[CurrentSelection]);
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
                    choice = TemplateOptions[CurrentSelection];
                    break;
                case ConsoleKey.RightArrow:
                    choice = TemplateOptions[CurrentSelection];
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
                    if (CurrentSelection + 1 <= TemplateOptions.Length - 1)
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
