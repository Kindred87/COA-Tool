using System;
using System.Collections.Generic;
using System.Text;

namespace COA_Tool.Console
{
    class Util
    {
        public Util()
        {
            SetSize();
        }
        private void SetSize()
        {
            System.Console.WindowWidth = (int)((double)System.Console.LargestWindowWidth * 0.75);
            System.Console.WindowHeight = (int)((double)System.Console.LargestWindowHeight * 0.5);
        }
        public static void WriteMessageInCenter(string message)
        {
            int cursorRow = (int)((double)System.Console.WindowHeight * 0.5);
            RemoveMessageInCenter();

            int cursorColumn = (int)((double)System.Console.WindowWidth * 0.5 - message.Length);

            System.Console.SetCursorPosition(cursorColumn, cursorRow);
            System.Console.WriteLine(message);
        }
        public static void RemoveMessageInCenter()
        {
            int cursorRow = (int)((double)System.Console.WindowHeight * 0.5);
            System.Console.SetCursorPosition(0, cursorRow);
            System.Console.Write(new string(' ', System.Console.WindowWidth));
        }
    }
}
