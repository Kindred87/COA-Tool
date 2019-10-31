using System;
using System.Collections.Generic;
using System.Text;
using System.Threading;

namespace CoA_Tool.Excel
{
    /// <summary>
    /// Handles operations necessary for the template's algorithm before initiating document generation
    /// </summary>
    class WorkbookByAlgorithm
    {
        public WorkbookByAlgorithm(Template template)
        {
            // Objects are instantiated with minimum processing until pertinent loading methods are called for the selected algorithm
            CSV.NWAData nwaData = new CSV.NWAData();
            CSV.Tableau tableau = new CSV.Tableau();
            FinishedGoods finishedGoods = new FinishedGoods();

            // For each sales order in tableau, spawn a workbook
            if (template.SelectedAlgorithm == Template.Algorithm.Standard)
            {
                nwaData.LoadCSVFiles();
                tableau.Load();
                finishedGoods.Load();

                foreach(List<List<string>> order in tableau.FileContents)
                {
                    Excel.Workbook workbook = new Workbook(template);
                    workbook.TableauData = order;
                    workbook.FinishedGoods = finishedGoods.Contents;
                    workbook.MicroResults = nwaData.DelimitedMicroResults;
                    workbook.TitrationResults = nwaData.DelimitedTitrationResults;

                    Thread thread = new Thread(workbook.Generate);
                    thread.Start();
                }
            }
            // For each group of relevant items that fall within the user-requested time span, spawn a workbook
            else if(template.SelectedAlgorithm == Template.Algorithm.DaysFromToday)
            {
                nwaData.LoadCSVFiles();
            }
            

            


            
        }
    }
}
