using System;
using System.Collections.Generic;
using System.Text;

namespace CoA_Tool.Console
{
    /// <summary>
    /// Contains miscellaneous methods for console operations
    /// </summary>
    class Util
    {
        // Constructor
        public Util()
        {
            
        }

        // Public methods
        /// <summary>
        /// Resizes the console window
        /// </summary>
        public static void SetSize()
        {
            System.Console.WindowWidth = (int)((double)System.Console.LargestWindowWidth * 0.75);
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
            int cursorRow = (int)((double)System.Console.WindowHeight * 0.5);
            RemoveMessageInCenter();

            int cursorColumn = (int)((double)(System.Console.WindowWidth - message.Length) * 0.5);

            System.Console.SetCursorPosition(cursorColumn, cursorRow);
            System.Console.WriteLine(message);
        }
        /// <summary>
        /// Writes a string in the center of the console window, color is reset to gray after writing
        /// </summary>
        /// <param name="message"></param>
        public static void WriteMessageInCenter(string message, System.ConsoleColor color)
        {
            int cursorRow = (int)((double)System.Console.WindowHeight * 0.5);
            RemoveMessageInCenter();

            int cursorColumn = (int)((double)(System.Console.WindowWidth - message.Length) * 0.5);

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
            int cursorRow = (int)((double)System.Console.WindowHeight * 0.5);
            System.Console.SetCursorPosition(0, cursorRow);
            System.Console.Write(new string(' ', System.Console.WindowWidth));
        }

        // Private methods
    }
}
