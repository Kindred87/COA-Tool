using System;
using System.Collections.Generic;
using System.Text;

namespace CoA_Tool.Excel
{
    /// <summary>
    /// Handles operations necessary for the template's algorithm before initiating document generation
    /// </summary>
    class WorkbookByAlgorithm
    {
        //Objects
        CSV.NWAData NWAData;
        public WorkbookByAlgorithm(Template template)
        {
            NWAData = new CSV.NWAData();
            CSV.Tableau tableau = new CSV.Tableau();
            FinishedGoods finishedGoods = new FinishedGoods();

            if (template.SelectedAlgorithm == Template.Algorithm.Standard)
            {
                NWAData.LoadCSVFiles();
                tableau.Load();
                finishedGoods.Load();
            }
            else if(template.SelectedAlgorithm == Template.Algorithm.DaysFromToday)
            {
                NWAData.LoadCSVFiles();
            }
            

            


            
        }
    }
}
