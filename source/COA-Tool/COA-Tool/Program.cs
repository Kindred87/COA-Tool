using System;
using System.Collections.Generic;
using COA_Tool;
using System.Threading;
using System.Threading.Tasks;

namespace COA_Tool
{
    class Program
    {
        static void Main(string[] args)
        {
            Console.Util util = new Console.Util();
            CSV.RequiredFiles requiredFiles = new CSV.RequiredFiles();
            System.Console.SetCursorPosition(0, 30);
            CSV.Tableau tableau = new CSV.Tableau();

            System.Console.Write("Customer: ");

            //ConsoleKey key = System.Console.ReadKey().Key;

            ConsoleKey key = System.Console.ReadKey().Key;

            System.Console.CursorTop = System.Console.CursorTop - 1;

            System.Console.Write(new string(' ', System.Console.WindowWidth));

            List<Task> tasks = new List<Task>();

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

                Excel.Workbook workbook = new Excel.Workbook(order, requiredFiles.DelimitedFinishedGoods, requiredFiles.DelimitedTitrationResults, requiredFiles.DelimitedMicroResults, customer);
                new Thread(new ThreadStart(workbook.Generate)).Start();
            }


        }
        public void CreateWorkbook(List<List<string>> order, List<List<string>> finishedGoods, List<List<string>> titrationResults, List<List<string>> microResults, Excel.Workbook.CustomerName customerName)
        {
            new Excel.Workbook(order, finishedGoods, titrationResults, microResults, customerName);
        }
    }
}



