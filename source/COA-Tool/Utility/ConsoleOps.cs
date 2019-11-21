using System;
using System.Collections.Generic;
using System.Text;

namespace CoA_Tool.Utility
{
    /// <summary>
    /// Contains miscellaneous methods for console operations
    /// </summary>
    class ConsoleOps // TODO: Move to CoA_Tool.Utility
    {
        // Constructor
        public ConsoleOps()
        {
            
        }

        // Public methods
        /// <summary>
        /// Resizes the console window
        /// </summary>
        public static void SetInitialSize()
        {
            System.Console.WindowWidth = (int)((double)System.Console.LargestWindowWidth * 0.8);
            System.Console.WindowHeight = (int)((double)System.Console.LargestWindowHeight * 0.5);
        }
        /// <summary>
        /// Sets the title for the console window
        /// </summary>
        public static void SetTitle()
        {
            System.Console.Title = "CoA Tool";
        }
        /// <summary>
        /// Writes a string in the center of the console window
        /// </summary>
        /// <param name="message"></param>
        public static void WriteMessageInCenter(string message)
        {
            int cursorRow = (int)(System.Console.WindowHeight * 0.5);
            RemoveMessageInCenter();

            int cursorColumn = (int)((System.Console.WindowWidth - message.Length) * 0.5);

            System.Console.SetCursorPosition(cursorColumn, cursorRow);
            System.Console.WriteLine(message);
        }
        /// <summary>
        /// Writes a string in the center of the console window, color is reset to gray after writing
        /// </summary>
        /// <param name="message"></param>
        public static void WriteMessageInCenter(string message, System.ConsoleColor color)
        {
            int cursorRow = (int)(System.Console.WindowHeight * 0.5);
            RemoveMessageInCenter();

            int cursorColumn = (int)((System.Console.WindowWidth - message.Length) * 0.5);

            System.Console.SetCursorPosition(cursorColumn, cursorRow);
            System.Console.ForegroundColor = color;
            System.Console.WriteLine(message);
            System.Console.ForegroundColor = ConsoleColor.Gray;
        }
        
        /// <summary>
        /// Overwrites any string in the "center" of the console window with whitespace
        /// </summary>
        public static void RemoveMessageInCenter()
        {
            int cursorRow = (int)(System.Console.WindowHeight * 0.5);
            System.Console.SetCursorPosition(0, cursorRow);
            System.Console.Write(new string(' ', System.Console.WindowWidth));
        }
        /// <summary>
        /// Converts a user's input to a DateTime object
        /// </summary>
        /// <param name="inputPrompt">The message prompt</param>
        /// <returns></returns>
        public static DateTime GetDateFromUser(string inputPrompt)
        {
            WriteMessageInCenter(inputPrompt);

            string userInput;
            DateTime inputAsDateTime;
            int cursorRow = (int)(System.Console.WindowHeight * 0.5 + 1);
            int cursorColumn = (int)(System.Console.WindowWidth * 0.5);
            System.Console.CursorVisible = true;

            do
            {
                System.Console.SetCursorPosition(cursorColumn, cursorRow);
                userInput = System.Console.ReadLine();
                System.Console.SetCursorPosition(0, cursorRow);
                System.Console.Write(new string(' ', System.Console.WindowWidth));
            } while (DateTime.TryParse(userInput, out inputAsDateTime) == false);

            RemoveMessageInCenter();
            System.Console.CursorVisible = false;

            return inputAsDateTime;
        }

        // Private methods
    }
}
