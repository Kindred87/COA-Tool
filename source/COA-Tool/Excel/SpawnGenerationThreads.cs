﻿using System;
using System.Collections.Generic;
using System.Text;
using System.Threading;

namespace CoA_Tool.Excel
{
    /// <summary>
    /// Handles operations necessary for the template's algorithm before initiating document generation
    /// </summary>
    static class SpawnGenerationThreads
    {
        public static void Go(Templates.Template template)
        {
            // Objects are instantiated with minimum processing until pertinent loading methods are called for the selected algorithm
            CSV.NWAData nwaData = new CSV.NWAData();
            CSV.TableauData tableau = new CSV.TableauData();
            FinishedGoods finishedGoods = new FinishedGoods();
            int numberOfDocumentsToGenerate = 0;

            // For each sales order in tableau, spawn a workbook
            if (template.SelectedAlgorithm == Templates.Template.Algorithm.Standard)
            {
                nwaData.LoadCSVFiles();
                tableau.Load();
                finishedGoods.Load();

                foreach (CSV.SalesOrder salesOrder in tableau.SalesOrders)
                {
                    Console.Util.WriteMessageInCenter("Generating " + ++numberOfDocumentsToGenerate + " CoA documents");
                    WorkbookData workbook = new WorkbookData(template)
                    {
                        SalesOrder = salesOrder,
                        FinishedGoods = finishedGoods.Contents,
                        MicroResults = nwaData.DelimitedMicroResults,
                        TitrationResults = nwaData.DelimitedTitrationResults
                    };

                    Thread thread = new Thread(workbook.Generate);
                    thread.Start();
                }
            }
            // For each group of relevant items that fall within the user-requested time span, spawn a workbook
            else if (template.SelectedAlgorithm == Templates.Template.Algorithm.FromDateOnwards)
            {
                nwaData.LoadCSVFiles();
                DateTime desiredStartDate = Console.Util.GetDateFromUser("Please enter a start date for the search algorithm.");

                WorkbookData workbook = new WorkbookData(template)
                {
                    MicroResults = nwaData.DelimitedMicroResults,
                    TitrationResults = nwaData.DelimitedTitrationResults,
                    StartDate = desiredStartDate
                };

                Thread thread = new Thread(workbook.Generate);
                thread.Start();
            }
        }
    }
}
