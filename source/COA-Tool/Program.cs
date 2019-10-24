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
            Console.Util.SetSize();
            Console.Util.SetTitle();
            System.Console.CursorVisible = false; // Only visible for text input

            Template template = new Template();

            Excel.WorkbookByAlgorithm byAlgorithm = new Excel.WorkbookByAlgorithm(template);

        }
        /*static void CreateExternalCOA(List<List<List<string>>> tableauData, CSV.Common common, Excel.FinishedGoods finishedGoods)
        {
            foreach(List<List<string>> order in tableauData)
            {
                Excel.Workbook workbook = new Excel.Workbook(common.DelimitedTitrationResults, common.DelimitedMicroResults, finishedGoods.Contents);
                workbook.TableauData = order;

                Thread thread = new Thread(workbook.Generate);
                thread.Start();
            }
        }
        static void CreateInternalCOA(CSV.Common common, Excel.FinishedGoods finishedGoods)
        {
            string input;
            int daysBackToInclude;

            do
            {
                System.Console.SetCursorPosition(0, 30);
                System.Console.Write(new string(' ', System.Console.WindowWidth));
                
                System.Console.SetCursorPosition(0, 30);
                System.Console.Write("Days before " + DateTime.Now.ToShortDateString() + ": ");
                
                input = System.Console.ReadLine();
            } while (int.TryParse(input, out daysBackToInclude) == false);

            // A string is used as the hashset can't easily differentiate arrays with matching contents
            HashSet<string> set = new HashSet<string>();

            foreach(List<string> line in common.DelimitedMicroResults)
            {
                for (int i = 0; i < daysBackToInclude; i++)
                {
                    if(line[5] == DateTime.Now.AddDays(i * -1).ToString("M/d/yy") || line[5] == DateTime.Now.AddDays(i * -1).ToShortDateString())
                    {
                        if(line[16] == "K1")
                        {
                            if(line[10].Contains('/') || line[10].Contains('\\'))
                            {
                                char[] delimit = { '/', '\\' };
                                string reformattedEntry = line[10].Split(delimit)[0];
                                reformattedEntry += " & ";
                                reformattedEntry += line[10].Split(delimit)[1];

                                set.Add(Convert.ToDateTime(line[9]).ToString("M-d-yy") + "," + reformattedEntry);
                            }
                            else
                            {
                                set.Add(Convert.ToDateTime(line[9]).ToString("M-d-yy") + "," + line[10]); 
                            }

                        }
                    }
                }
            }

            foreach(string productAndDateCombo in set)
            {
                Excel.Workbook workbook = new Excel.Workbook(common.DelimitedTitrationResults, common.DelimitedMicroResults, finishedGoods.Contents, common.Recipes);
                workbook.InternalCOAData = productAndDateCombo.Split(new char[] { ',' });

                Thread thread = new Thread(workbook.Generate);
                thread.Start();
            }
        }*/
    }
}



