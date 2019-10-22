using System;
using System.Collections.Generic;
using System.Text;

namespace CoA_Tool.Excel
{
    class WorkbookByAlgorithm
    {
        public WorkbookByAlgorithm(Template template)
        {
            if(template.SelectedAlgorithm == Template.Algorithm.Standard)
            {

            }
            else if(template.SelectedAlgorithm == Template.Algorithm.DaysFromToday)
            {

            }
            Excel.FinishedGoods finishedGoods = new Excel.FinishedGoods();

            CSV.Common requiredFiles = new CSV.Common();

            if (requiredFiles.AllFilesReady == false)
            {
                throw new Exception("File loading error");
            }

            CSV.Tableau tableau = new CSV.Tableau();
        }
    }
}
