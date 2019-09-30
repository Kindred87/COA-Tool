using System;
using System.Collections.Generic;
using System.Threading;

namespace CoA_Tool
{
    class Program
    {
        static void Main(string[] args)
        {
            Console.Util.SetSize();
            Console.Util.SetTitle();

            Excel.FinishedGoods finishedGoods = new Excel.FinishedGoods();

            CSV.Common requiredFiles = new CSV.Common();

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
                    customer = Excel.Workbook.CustomerName.Latitude36;
                }

                Excel.Workbook workbook = new Excel.Workbook(order, requiredFiles.DelimitedTitrationResults, requiredFiles.DelimitedMicroResults, customer, finishedGoods.Contents, requiredFiles.Recipes);

                Thread thread = new Thread(workbook.Generate);
                thread.Start();
            }
        }
    }
}



