using System;
using System.Collections.Generic;
using System.Text;

namespace CoA_Tool.Excel
{
    class WorkbookByAlgorithm
    {
        //Objects
        CSV.Common RequiredFiles;
        public WorkbookByAlgorithm(Template template)
        {
            RequiredFiles = new CSV.Common();

            if (template.SelectedAlgorithm == Template.Algorithm.Standard)
            {
                RequiredFiles.LoadCSVFiles();
            }
            else if(template.SelectedAlgorithm == Template.Algorithm.DaysFromToday)
            {

            }
            Excel.FinishedGoods finishedGoods = new Excel.FinishedGoods();

            


            CSV.Tableau tableau = new CSV.Tableau();
        }
    }
}
