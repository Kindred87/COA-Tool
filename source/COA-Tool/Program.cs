using System;
using System.Collections.Generic;
using System.Threading;
using CoA_Tool;

namespace CoA_Tool
{
    class Program
    {
        static void Main(string[] args)
        {
            Utility.ConsoleOps.SetInitialSize();
            Utility.ConsoleOps.SetTitle();
            System.Console.CursorVisible = false;

            Excel.SpawnGenerationThreads.Go(new Templates.Template());
        }
    }
}



