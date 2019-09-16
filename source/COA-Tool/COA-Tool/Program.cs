using System;
using System.Collections.Generic;
using COA_Tool;
using System.Threading;

namespace COA_Tool
{
    class Program
    {
        static void Main(string[] args)
        {
            Console.Util util = new Console.Util();
            CSV.RequiredFiles requiredFiles = new CSV.RequiredFiles();
            if(requiredFiles.AllFilesReady == false)
            {
                throw new Exception("File loading error");
            }
            System.Console.SetCursorPosition(0, 30);
            CSV.Tableau tableau = new CSV.Tableau();

            System.Console.Write("Customer: ");

            ConsoleKey key = System.Console.ReadKey().Key;

            System.Console.CursorTop = System.Console.CursorTop - 1;

            System.Console.Write(new string(' ', System.Console.WindowWidth));

            foreach (List<List<string>> order in tableau.FileContents)
            {
                Excel.Workbook.CustomerName customer = Excel.Workbook.CustomerName.TaylorFarmsTennessee;

                if (key == ConsoleKey.D1)
                {
                    customer = Excel.Workbook.CustomerName.TaylorFarmsTennessee;
                }
                else if (key == System.ConsoleKey.D2)
                {
                    customer = Excel.Workbook.CustomerName.Latitute36;
                }

                Excel.Workbook workbook = new Excel.Workbook(order, requiredFiles.DelimitedTitrationResults, requiredFiles.DelimitedMicroResults, customer, requiredFiles.FinishedGoods, requiredFiles.Recipes);

                Thread thread = new Thread(workbook.Generate);
                thread.Start();
            }
        }
    }
}



