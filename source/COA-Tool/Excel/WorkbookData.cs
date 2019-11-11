using System;
using System.Collections.Generic;
using System.Text;
using OfficeOpenXml;
using System.Xml;
using System.Drawing;
using OfficeOpenXml.Style;
using System.IO;
using OfficeOpenXml.Drawing;


namespace CoA_Tool.Excel
{
    /// <summary>
    /// Main document generation class, end of program flow
    /// </summary>
    class WorkbookData
    {
        // Enums
        private enum TitrationOffset { Acidity = 5, Viscosity = 2, Salt = 4, pH = 6 }
        private enum MicroOffset { Yeast = 9, Mold = 11, Aerobic = 15, Coliform = 7, Lactic = 13, EColiform = 5  }

        //  Strings
        private string AcidMethod = "(AOAC 30.048 14th Ed.)  ";
        private string pHMethod = "(AOAC 30.012 14th Ed.)  ";
        private string ViscosityCMMethod = "(Bostwick)  ";
        private string ViscosityCPSMethod = "(Brookfield)  ";
        private string SaltMethod = "(AOAC 937.09 18th Ed.)  ";
        private string YeastMethod = "(AOAC 997.02)  ";
        private string MoldMethod = "(AOAC 997.02)  ";
        private string AerobicMethod = "(AOAC 990.12)  ";
        private string ColiformMethod = "(AOAC 991.14)  ";
        private string EColiMethod = "(AOAC 991.14)  ";
        private string LacticMethod = "(AOAC 990.12)  ";

        public string[] InternalCOAData; // {made date, item code}

        // DateTimes
        public DateTime StartDate;

        //  Bools
        private bool SaveFile = false;

        //  Lists
        public List<List<string>> TableauData;
        public List<List<string>> TitrationResults;
        public List<List<string>> MicroResults;
        public List<List<string>> FinishedGoods;

        private List<List<string>> RecipeAndItemValuesFromFilterCheck;

        // Hashsets
        private HashSet<string> BatchIndicesToIgnore;

        //  Objects
        private Templates.Template WorkbookTemplate;
        public CSV.SalesOrder SalesOrder { get; set; }

        // Constructor
        public WorkbookData(Templates.Template template)
        {
            WorkbookTemplate = template;
        }

        public void Generate()
        {
            using (ExcelPackage package = new ExcelPackage())
            {
                int pageCount = 0;

                if(WorkbookTemplate.SelectedAlgorithm == Templates.Template.Algorithm.Standard)
                {
                    pageCount = (SalesOrder.Lots.Count- 1) / 6;
                    if ((SalesOrder.Lots.Count - 1) % 6 > 0)
                    {
                        pageCount++;
                    }
                }
                else if(WorkbookTemplate.SelectedAlgorithm == Templates.Template.Algorithm.FromDateOnwards)
                {
                    int itemCount = 0;

                    foreach(List<string> line in MicroResults)
                    {
                        if (line[9] == Convert.ToDateTime(InternalCOAData[0]).ToString("M/d/yy") && line[10] == InternalCOAData[1])
                            itemCount++;
                    }
                        pageCount = itemCount / 6;
                        if ((itemCount) % 6 > 0)
                            pageCount++;
                }

                for (int i = 1; i <= pageCount; i++)
                {
                    ExcelWorksheet worksheet = package.Workbook.Worksheets.Add("Page " + i);
                    new Worksheet.Page(WorkbookTemplate, worksheet);
                    SaveFile = true;
                }

                if(SaveFile)
                {
                    if (Directory.Exists(Environment.GetFolderPath(Environment.SpecialFolder.Desktop) + "/CoAs") == false)
                        Directory.CreateDirectory(Environment.GetFolderPath(Environment.SpecialFolder.Desktop) + "/CoAs");

                    if (Directory.Exists(Environment.GetFolderPath(Environment.SpecialFolder.Desktop) + "/CoAs/Internal") == false)
                        Directory.CreateDirectory(Environment.GetFolderPath(Environment.SpecialFolder.Desktop) + "/CoAs/Internal");

                    if (WorkbookTemplate.SelectedAlgorithm == Templates.Template.Algorithm.Standard)
                        package.SaveAs(new FileInfo(Environment.GetFolderPath(Environment.SpecialFolder.Desktop) + "/CoAs/" + SalesOrder.OrderNumber + ".xlsx"));
                    else
                        package.SaveAs(new FileInfo(Environment.GetFolderPath(Environment.SpecialFolder.Desktop) + "/CoAs/Internal/" + InternalCOAData[1] +
                            " (" + InternalCOAData[0] + ")" + ".xlsx")); 
                }
                else
                {
                    // Generally only triggered with internal CoAs with no available micro results
                }
            }
        }
        /// <summary>
        /// Populates content for the worksheet
        /// </summary>
        /// <param name="targetWorksheet"></param>
        private void PopulateWorksheetContents(ExcelWorksheet targetWorksheet, int currentPage, out bool WorkOnNextWorksheet)
        {
            WorkOnNextWorksheet = true;

            targetWorksheet.Cells.Style.Font.SetFromFont(new Font("Calibri", 11));

            targetWorksheet.View.ShowGridLines = false;

            targetWorksheet.Cells["C11:H30"].Style.HorizontalAlignment = ExcelHorizontalAlignment.Center;
            targetWorksheet.Cells["C11:H30"].Style.VerticalAlignment = ExcelVerticalAlignment.Center;

            targetWorksheet.Column(1).Width = 10.85;
            targetWorksheet.Column(2).Width = 11.4;
            targetWorksheet.Column(3).Width = 15.72;
            targetWorksheet.Column(4).Width = 15.72;
            targetWorksheet.Column(5).Width = 15.72;
            targetWorksheet.Column(6).Width = 15.72;
            targetWorksheet.Column(7).Width = 15.72;
            targetWorksheet.Column(8).Width = 15.72;

            targetWorksheet.Row(1).Height = 0;

            Image image = Image.FromFile("LH logo.png");
            ExcelPicture logo = targetWorksheet.Drawings.AddPicture("Logo", image);
            logo.SetSize(234, 124);
            logo.SetPosition(0, 301);

            targetWorksheet.Cells["A8"].Value = "Certificate of Analysis";
            targetWorksheet.Cells["A8"].Style.Font.SetFromFont(new Font("Calibri", 26, FontStyle.Bold));
            targetWorksheet.Cells["A8"].Style.HorizontalAlignment = ExcelHorizontalAlignment.Center;
            targetWorksheet.Cells["A8:H9"].Merge = true;

            targetWorksheet.Cells[11, 1, 60, 2].Style.Font.Size = 9;
            targetWorksheet.Cells[11, 1, 60, 8].Style.VerticalAlignment = ExcelVerticalAlignment.Center;

            
            
            // Counters are used since document contents/placement is dynamic
            int currentRow = 11;
            int sizeOfFirstContentBlock = 0;


            // For first content block
            if (WorkbookTemplate.IncludeCustomerName)
            {
                sizeOfFirstContentBlock++;
                
                targetWorksheet.Cells[currentRow, 1, currentRow, 2].Merge = true;
                targetWorksheet.Cells[currentRow, 1].Value = "Customer";
                targetWorksheet.Cells[currentRow, 3, currentRow, 8].Merge = true;
                targetWorksheet.Cells[currentRow, 3].Value = WorkbookTemplate.Menu.UserChoice;

                currentRow++;

            } 

            if(WorkbookTemplate.IncludeSalesOrder)
            {
                sizeOfFirstContentBlock++;

                targetWorksheet.Cells[currentRow, 1, currentRow, 2].Merge = true;
                targetWorksheet.Cells[currentRow, 1].Value = "Customer S/O #";
                targetWorksheet.Cells[currentRow, 3, currentRow, 8].Merge = true;
                targetWorksheet.Cells[currentRow, 3].Value = TableauData[1][3];
                targetWorksheet.Cells[currentRow, 3].Style.Numberformat.Format = "0";

                currentRow++;
            }

            if(WorkbookTemplate.IncludePurchaseOrder)
            {
                sizeOfFirstContentBlock++;

                targetWorksheet.Cells[currentRow, 1, currentRow, 2].Merge = true;
                targetWorksheet.Cells[currentRow, 1].Value = "PO #";
                targetWorksheet.Cells[currentRow, 3, currentRow, 8].Merge = true;
                targetWorksheet.Cells[currentRow, 3].Style.Numberformat.Format = "0";

                currentRow++;
            }

            if(WorkbookTemplate.IncludeGenerationDate)
            {
                sizeOfFirstContentBlock++;

                targetWorksheet.Cells[currentRow, 1, currentRow, 2].Merge = true;
                targetWorksheet.Cells[currentRow, 1].Value = "Date";
                targetWorksheet.Cells[currentRow, 3, currentRow, 8].Merge = true;
                targetWorksheet.Cells[currentRow, 3].Value = TableauData[1][3];
                targetWorksheet.Cells[currentRow, 3].Style.Numberformat.Format = "m/d/yyyy";
                targetWorksheet.Cells[currentRow, 3].Value = DateTime.Now.Date.ToShortDateString();

                currentRow++;

            }

            // For first content block
            if(sizeOfFirstContentBlock > 0)
            {
                targetWorksheet.Cells[10, 3, 10, 8].Style.Border.Bottom.Style = ExcelBorderStyle.Medium;
                targetWorksheet.Cells[11, 3, 10 + sizeOfFirstContentBlock, 8].Style.Font.Size = 12;
                targetWorksheet.Cells[11, 3, 10 + sizeOfFirstContentBlock, 8].Style.Font.Bold = true;
                targetWorksheet.Cells[11, 3, 10 + sizeOfFirstContentBlock, 8].Style.HorizontalAlignment = ExcelHorizontalAlignment.Center;

                CreateTableOfBorders(6, sizeOfFirstContentBlock, 11, 3, targetWorksheet);
            }

            // For second content block

            List<string> lotsToProcess = new List<string>();
            for (int i = 6 * (currentPage - 1); i < 6 * currentPage; i++)
            {
                string lot = GetLotCode(i + 1);

                if (lot != string.Empty)
                    lotsToProcess.Add(lot);
            }

            if (sizeOfFirstContentBlock > 0)
                currentRow += 2; // An empty space between blocks and an empty space for the block's header if the block size > 0
            else
                currentRow = 12; // Allows the second block's header to use row 11

            int sizeOfSecondContentBlock = 0;

            if (WorkbookTemplate.IncludeProductName)
            {
                sizeOfSecondContentBlock++;

                targetWorksheet.Cells[currentRow, 1, currentRow, 2].Merge = true;
                targetWorksheet.Cells[currentRow, 1].Value = "Product";
                targetWorksheet.Cells[currentRow, 3, currentRow, 8].Style.WrapText = true;

                for(int i = 0; i < lotsToProcess.Count; i++)
                {
                    targetWorksheet.Cells[currentRow, 3 + i].Value = GetProductName(GetProductCode(lotsToProcess[i]));
                }

                currentRow++;
            }

            if(WorkbookTemplate.IncludeRecipeAndItem)
            {
                sizeOfSecondContentBlock++;

                targetWorksheet.Cells[currentRow, 1, currentRow, 2].Merge = true;
                targetWorksheet.Cells[currentRow, 1].Value = "Recipe/Item";

                for(int i = 0; i < lotsToProcess.Count; i++)
                {
                    string productCode = GetProductCode(lotsToProcess[i]);
                    string writeToCell = GetRecipeCode(productCode) +
                        "/" + productCode;
                    if(WorkbookTemplate.CustomFilters.Count > 0)
                    {
                        foreach (Templates.CustomFilter filter in WorkbookTemplate.CustomFilters)
                        {
                            if(filter.IsValidFilter && filter.ContentItem == Templates.Template.ContentItems.RecipeAndItem)
                            {
                                if(filter.FilterType == Templates.CustomFilter.FilterTypes.Whitelist)
                                {
                                    foreach (string criteria in filter.Criteria)
                                    {
                                        if(criteria != writeToCell)
                                        {
                                            return;
                                        }
                                    }
                                }
                                else // filter.FilterType == FilterTypes.Blacklist
                                {
                                    foreach(string criteria in filter.Criteria)
                                    {
                                        if(criteria == writeToCell)
                                        {
                                            return;
                                        }
                                    }
                                }
                            }
                        }
                    }
                    targetWorksheet.Cells[currentRow, 3 + i].Value = writeToCell;
                }

                currentRow++;
            }

            if(WorkbookTemplate.IncludeLotCode)
            {
                sizeOfSecondContentBlock++;

                targetWorksheet.Cells[currentRow, 1, currentRow, 2].Merge = true;
                targetWorksheet.Cells[currentRow, 3, currentRow, 8].Style.Numberformat.Format = "0";
                targetWorksheet.Cells[currentRow, 1].Value = "Lot Code";
                
                for (int i = 0; i < lotsToProcess.Count; i++)
                {
                    long convertedLotValue;

                    if (Int64.TryParse(lotsToProcess[i], out convertedLotValue) && lotsToProcess[i].Length == 13)
                    {
                        targetWorksheet.Cells[currentRow, 3 + i].Value = convertedLotValue;
                    }
                    else if(lotsToProcess[i].Length > 13)
                    {
                        targetWorksheet.Cells[currentRow, 3 + i].Value = "Lot too long";
                        targetWorksheet.Cells[currentRow, 3 + i].Style.Font.Color.SetColor(Color.Red);
                    }
                    else if(lotsToProcess[i].Length < 13)
                    {
                        targetWorksheet.Cells[currentRow, 3 + i].Value = "Lot too short";
                        targetWorksheet.Cells[currentRow, 3 + i].Style.Font.Color.SetColor(Color.Red);
                    }
                    else
                    {
                        targetWorksheet.Cells[currentRow, 3 + i].Value = "Invalid lot";
                        targetWorksheet.Cells[currentRow, 3 + i].Style.Font.Color.SetColor(Color.Red);
                    }
                }
                currentRow++;
            }

            if(WorkbookTemplate.IncludeBatch)
            {
                sizeOfSecondContentBlock++;

                targetWorksheet.Cells[currentRow, 1, currentRow, 2].Merge = true;
                targetWorksheet.Cells[currentRow, 1].Value = "Batch";

                if(WorkbookTemplate.SelectedAlgorithm == Templates.Template.Algorithm.FromDateOnwards)
                {
                    for (int i = 0; i < lotsToProcess.Count; i++)
                    {
                        targetWorksheet.Cells[currentRow, 3 + i].Value = "";
                    }
                }
                currentRow++;
            }

            if(WorkbookTemplate.IncludeBestByDate)
            {
                sizeOfSecondContentBlock++;

                targetWorksheet.Cells[currentRow, 1, currentRow, 2].Merge = true;
                targetWorksheet.Cells[currentRow, 1].Value = "Best By Date";

                currentRow++;
            }

            if(WorkbookTemplate.IncludeManufacturingDate)
            {
                sizeOfSecondContentBlock++;

                targetWorksheet.Cells[currentRow, 1, currentRow, 2].Merge = true;
                targetWorksheet.Cells[currentRow, 1].Value = "Manufacturing Date";

                currentRow++;
            }

            if(WorkbookTemplate.IncludeManufacturingSite)
            {
                sizeOfSecondContentBlock++;

                targetWorksheet.Cells[currentRow, 1, currentRow, 2].Merge = true;
                targetWorksheet.Cells[currentRow, 1].Value = "Manufacturing Site";

                currentRow++;
            }

            // For both first and second content blocks: left-hand cell descriptions, second content block header, border drawing

            int sumOfBlockRows = sizeOfFirstContentBlock + sizeOfSecondContentBlock;

            if (sumOfBlockRows > 0)
            {
                if(sizeOfSecondContentBlock > 0)
                {
                    // Writes header for second content block
                    int secondBlockHeaderRow;
                    if (sizeOfFirstContentBlock == 0)
                    {
                        secondBlockHeaderRow = 11;
                        sumOfBlockRows++; // Covers second block's header row
                    }
                    else
                    {
                        sumOfBlockRows += 2; // Covers empty rows separating segments
                        secondBlockHeaderRow = 11 + sizeOfFirstContentBlock + 1; // 11 is starting row, 1 is empty space between blocks
                    }

                    targetWorksheet.Cells[secondBlockHeaderRow, 3].Value = "Product Information";
                    targetWorksheet.Cells[secondBlockHeaderRow, 3, secondBlockHeaderRow, 8].Merge = true;
                    targetWorksheet.Cells[secondBlockHeaderRow, 3].Style.Font.Bold = true;
                    targetWorksheet.Cells[secondBlockHeaderRow, 3, secondBlockHeaderRow, 8].Style.Border.Bottom.Style = ExcelBorderStyle.Medium;
                    targetWorksheet.Cells[secondBlockHeaderRow, 3, secondBlockHeaderRow + sizeOfSecondContentBlock, 
                        8].Style.HorizontalAlignment = ExcelHorizontalAlignment.Center;

                    CreateTableOfBorders(6, sizeOfSecondContentBlock, secondBlockHeaderRow + 1, 3, targetWorksheet);
                }

                targetWorksheet.Cells[11, 1, 10 + sumOfBlockRows, 2].Style.Font.Italic = true;
                targetWorksheet.Cells[11, 1, 10 + sumOfBlockRows, 2].Style.HorizontalAlignment = ExcelHorizontalAlignment.Right;
            }

            // For third content block

            if (sumOfBlockRows > 0)
                currentRow += 2; // An empty space between blocks and an empty space for the block's header if the total block size > 0
            else
                currentRow = 12; // Allows the third block's header to use row 11

            int sizeOfThirdContentBlock = 0;
            
            if(WorkbookTemplate.IncludeAcidity)
            {
                sizeOfThirdContentBlock++;

                targetWorksheet.Cells[currentRow, 1].Value = "% Acid (TA)";
                targetWorksheet.Cells[currentRow, 2].Value = AcidMethod;

                currentRow++;
            }

            if(WorkbookTemplate.IncludepH)
            {
                sizeOfThirdContentBlock++;

                targetWorksheet.Cells[currentRow, 1].Value = "pH";
                targetWorksheet.Cells[currentRow, 2].Value = pHMethod;

                currentRow++;
            }

            if(WorkbookTemplate.IncludeViscosityCM)
            {
                sizeOfThirdContentBlock++;

                targetWorksheet.Cells[currentRow, 1].Value = "Viscosity cm";
                targetWorksheet.Cells[currentRow, 2].Value = ViscosityCMMethod;

                currentRow++;
            }

            if(WorkbookTemplate.IncludeViscosityCPS)
            {
                sizeOfThirdContentBlock++;

                targetWorksheet.Cells[currentRow, 1].Value = "Viscosity cps";
                targetWorksheet.Cells[currentRow, 2].Value = ViscosityCPSMethod;

                currentRow++;
            }

            if(WorkbookTemplate.IncludeWaterActivity)
            {
                sizeOfThirdContentBlock++;

                targetWorksheet.Cells[currentRow, 1].Value = "Water activity (aW)";

                currentRow++;
            }

            if(WorkbookTemplate.IncludeBrixSlurry)
            {
                sizeOfThirdContentBlock++;

                targetWorksheet.Cells[currentRow, 1].Value = "Brix slurry";

                currentRow++;
            }

            if(sizeOfThirdContentBlock > 0)
            {
                sumOfBlockRows += sizeOfThirdContentBlock;

                int thirdBlockHeaderRow;

                if (sumOfBlockRows - sizeOfThirdContentBlock == 0)
                {
                    thirdBlockHeaderRow = 11;
                }
                else
                {
                    thirdBlockHeaderRow = 11 + sumOfBlockRows - sizeOfThirdContentBlock + 1; // 11 is starting row, 1 is empty space between blocks
                    sumOfBlockRows += 2;
                }

                CreateTableOfBorders(6, sizeOfThirdContentBlock, thirdBlockHeaderRow + 1, 3, targetWorksheet);

                targetWorksheet.Cells[thirdBlockHeaderRow, 1].Value = "Test";
                targetWorksheet.Cells[thirdBlockHeaderRow, 2].Value = "Method";
                targetWorksheet.Cells[thirdBlockHeaderRow, 1, thirdBlockHeaderRow, 3].Style.HorizontalAlignment = ExcelHorizontalAlignment.Center;
                targetWorksheet.Cells[thirdBlockHeaderRow, 3, thirdBlockHeaderRow + sizeOfThirdContentBlock, 
                    8].Style.HorizontalAlignment = ExcelHorizontalAlignment.Center;
                targetWorksheet.Cells[thirdBlockHeaderRow, 3].Value = "Analytical Results";

                targetWorksheet.Cells[thirdBlockHeaderRow, 3, thirdBlockHeaderRow, 8].Merge = true;
                targetWorksheet.Cells[thirdBlockHeaderRow, 1, thirdBlockHeaderRow, 3].Style.Font.Bold = true;
                targetWorksheet.Cells[thirdBlockHeaderRow, 1, thirdBlockHeaderRow, 8].Style.Border.Bottom.Style = ExcelBorderStyle.Medium;

                targetWorksheet.Cells[thirdBlockHeaderRow + 1, 2, thirdBlockHeaderRow + sizeOfThirdContentBlock, 2].Style.Font.Size = 6;
                targetWorksheet.Cells[thirdBlockHeaderRow + 1, 2, thirdBlockHeaderRow + sizeOfThirdContentBlock, 
                    2].Style.HorizontalAlignment = ExcelHorizontalAlignment.Right;
            }

            // For fourth content block

            if (sumOfBlockRows > 0)
                currentRow += 2; // An empty space between blocks and an empty space for the block's header if the total block size > 0
            else
                currentRow = 12; // Allows the third block's header to use row 11

            int sizeOfFourthContentBlock = 0;

            if(WorkbookTemplate.IncludeYeast)
            {
                sizeOfFourthContentBlock++;

                targetWorksheet.Cells[currentRow, 1].Value = "Yeast";
                targetWorksheet.Cells[currentRow, 2].Value = YeastMethod;

                currentRow++;
            }

            if(WorkbookTemplate.IncludeMold)
            {
                sizeOfFourthContentBlock++;

                targetWorksheet.Cells[currentRow, 1].Value = "Mold";
                targetWorksheet.Cells[currentRow, 2].Value = MoldMethod;

                currentRow++;
            }

            if(WorkbookTemplate.IncludeAerobic)
            {
                sizeOfFourthContentBlock++;

                targetWorksheet.Cells[currentRow, 1].Value = "Aerobic";
                targetWorksheet.Cells[currentRow, 2].Value = AerobicMethod;

                currentRow++;
            }

            if(WorkbookTemplate.IncludeColiform)
            {
                sizeOfFourthContentBlock++;

                targetWorksheet.Cells[currentRow, 1].Value = "Total coliform";
                targetWorksheet.Cells[currentRow, 2].Value = ColiformMethod;

                currentRow++;
            }

            if (WorkbookTemplate.IncludeEColi)
            {
                sizeOfFourthContentBlock++;

                targetWorksheet.Cells[currentRow, 1].Value = "E. coliform";
                targetWorksheet.Cells[currentRow, 2].Value = EColiMethod;

                currentRow++;
            }

            if(WorkbookTemplate.IncludeLactics)
            {
                sizeOfFourthContentBlock++;

                targetWorksheet.Cells[currentRow, 1].Value = "Lactics";
                targetWorksheet.Cells[currentRow, 2].Value = LacticMethod;

                currentRow++;
            }

            if(WorkbookTemplate.IncludeSalmonella)
            {
                sizeOfFourthContentBlock++;

                targetWorksheet.Cells[currentRow, 1].Value = "Salmonella";

                currentRow++;
            }

            if(WorkbookTemplate.IncludeListeria)
            {
                sizeOfFourthContentBlock++;

                targetWorksheet.Cells[currentRow, 1].Value = "Listeria";

                currentRow++;
            }

            if(sizeOfFourthContentBlock > 0)
            {
                sumOfBlockRows += sizeOfFourthContentBlock;

                int fourthBlockHeaderRow;

                if (sumOfBlockRows - sizeOfFourthContentBlock == 0)
                {
                    fourthBlockHeaderRow = 11;
                }
                else
                {
                    fourthBlockHeaderRow = 11 + sumOfBlockRows - sizeOfFourthContentBlock + 1; // 11 is starting row, 1 is empty space between blocks
                    sumOfBlockRows += 2;
                }

                targetWorksheet.Cells[fourthBlockHeaderRow, 1].Value = "Test (cfu/gram)";
                targetWorksheet.Cells[fourthBlockHeaderRow, 2].Value = "Method";
                targetWorksheet.Cells[fourthBlockHeaderRow, 3].Value = "Microbiological Results";
                targetWorksheet.Cells[fourthBlockHeaderRow, 1, fourthBlockHeaderRow, 3].Style.Font.Bold = true;
                targetWorksheet.Cells[fourthBlockHeaderRow, 1, fourthBlockHeaderRow, 3].Style.HorizontalAlignment = ExcelHorizontalAlignment.Center;
                targetWorksheet.Cells[fourthBlockHeaderRow, 3, fourthBlockHeaderRow + sizeOfFourthContentBlock,
                    8].Style.HorizontalAlignment = ExcelHorizontalAlignment.Center;
                targetWorksheet.Cells[fourthBlockHeaderRow, 3, fourthBlockHeaderRow, 8].Merge = true;
                targetWorksheet.Cells[fourthBlockHeaderRow, 1, fourthBlockHeaderRow, 8].Style.Border.Bottom.Style = ExcelBorderStyle.Medium;
                targetWorksheet.Cells[fourthBlockHeaderRow + 1, 2, fourthBlockHeaderRow + sizeOfFourthContentBlock, 2].Style.Font.Size = 6;
                targetWorksheet.Cells[fourthBlockHeaderRow + 1, 2, fourthBlockHeaderRow + sizeOfFourthContentBlock, 2].Style.HorizontalAlignment = ExcelHorizontalAlignment.Right;

                CreateTableOfBorders(6, sizeOfFourthContentBlock, fourthBlockHeaderRow + 1, 3, targetWorksheet);
            }

            // For fifth content block
            if (sumOfBlockRows > 0)
                currentRow += 2; // An empty space between blocks and an empty space for the block's header if the total block size > 0
            else
                currentRow = 12; // Allows the third block's header to use row 11

            int sizeOfFifthContentBlock = 0;

            if(WorkbookTemplate.IncludeColorAndAppearance)
            {
                sizeOfFifthContentBlock++;

                targetWorksheet.Cells[currentRow, 1].Value = "Color/Appearance";

                currentRow++;
            }

            if(WorkbookTemplate.IncludeForm)
            {
                sizeOfFifthContentBlock++;
                
                targetWorksheet.Cells[currentRow, 1].Value = "Form";

                currentRow++;
            }

            if(WorkbookTemplate.IncludeFlavorAndOdor)
            {
                sizeOfFifthContentBlock++;

                targetWorksheet.Cells[currentRow, 1].Value = "Flavor/Odor";
                
                currentRow++;
            }

            if(sizeOfFifthContentBlock > 0)
            {
                sumOfBlockRows += sizeOfFifthContentBlock;

                int fifthBlockHeaderRow;

                if (sumOfBlockRows - sizeOfFifthContentBlock == 0)
                {
                    fifthBlockHeaderRow = 11;
                }
                else
                {
                    fifthBlockHeaderRow = 11 + sumOfBlockRows - sizeOfFifthContentBlock + 1; // 11 is starting row, 1 is empty space between blocks
                    sumOfBlockRows += 2;
                }

                targetWorksheet.Cells[fifthBlockHeaderRow, 1].Value = "Test";
                targetWorksheet.Cells[fifthBlockHeaderRow, 3].Value = "Physical Characteristics";
                targetWorksheet.Cells[fifthBlockHeaderRow, 1, fifthBlockHeaderRow, 3].Style.Font.Bold = true;
                targetWorksheet.Cells[fifthBlockHeaderRow, 3, fifthBlockHeaderRow, 8].Merge = true;
                targetWorksheet.Cells[fifthBlockHeaderRow, 3, fifthBlockHeaderRow + sizeOfFifthContentBlock,
                    8].Style.HorizontalAlignment = ExcelHorizontalAlignment.Center;
                targetWorksheet.Cells[fifthBlockHeaderRow, 1, fifthBlockHeaderRow, 8].Style.Border.Bottom.Style = ExcelBorderStyle.Medium;

                CreateTableOfBorders(6, sizeOfFifthContentBlock, fifthBlockHeaderRow + 1, 3, targetWorksheet);
            }

            // For verified by and document disclaimer

            currentRow += 1;

            targetWorksheet.Cells[currentRow, 1].Value = "Verified By";
            targetWorksheet.Cells[currentRow, 1, currentRow + 5, 8].Style.Font.Size = 11;
            targetWorksheet.Cells[currentRow, 1].Style.Font.Bold = true;
            targetWorksheet.Cells[currentRow, 1, currentRow, 8].Merge = true;
            targetWorksheet.Cells[currentRow, 1, currentRow, 8].Style.Border.Bottom.Style = ExcelBorderStyle.Medium;

            currentRow++;
            targetWorksheet.Cells[currentRow, 1, currentRow, 8].Merge = true;
            CreateTableOfBorders(8, 1, currentRow, 1, targetWorksheet);
            targetWorksheet.Cells[currentRow - 1, 1, currentRow + 2, 8].Style.HorizontalAlignment = ExcelHorizontalAlignment.Center;

            currentRow += 2;
            targetWorksheet.Cells[currentRow, 1, currentRow + 4, 8].Merge = true;
            targetWorksheet.Cells[currentRow, 1].Style.WrapText = true;
            targetWorksheet.Cells[currentRow, 1].Value = "Please be advised that the results herein are accurate and representative to " +
                "the best of our knowledge, based on current data and information as of the date of this document.  This COA is intended " +
                "only for the person or company specified hereon as the recipient.  This COA shall not be distributed to any person or " +
                "company other than the intended recipient.  A person or company other than the one named hereon shall not rely or make " +
                "use of the statements and results herein. ";

            // For miscellaneous document items
            targetWorksheet.PrinterSettings.FitToPage = true;
            
            targetWorksheet.HeaderFooter.EvenFooter.LeftAlignedText = "created: 04/22/2016";
            targetWorksheet.HeaderFooter.EvenFooter.RightAlignedText = "10/31/2019 REV 03 F142-087";
            targetWorksheet.HeaderFooter.OddFooter.LeftAlignedText = "created: 04/22/2016";
            targetWorksheet.HeaderFooter.OddFooter.RightAlignedText = "10/31/2019 REV 03 F142-087";
            
            SaveFile = true;
        }
        /// <summary>
        /// Determines which customer-specific worksheet design to use and calls the appropriate method
        /// </summary>
        /// <param name="worksheet"></param>
        private void PopulateContentsByCustomer(ExcelWorksheet worksheet, int page)
        {
            PopulateContentsLatitude36(worksheet, page);
        }
        private void PopulateContentsTaylorFarmTennessee(ExcelWorksheet worksheet, int page)
        {
            SaveFile = true;

            CreateTableOfBorders(6, 4, 11, 3, worksheet);
            CreateTableOfBorders(6, 5, 16, 3, worksheet);
            CreateTableOfBorders(6, 4, 22, 3, worksheet);
            CreateTableOfBorders(6, 1, 27, 3, worksheet);

            worksheet.Cells["C11"].Value = "Taylor Farms / Tennessee";
            worksheet.Cells["C12"].Value = Convert.ToInt32(TableauData[1][3]);

            
            worksheet.Cells["A11"].Value = "Customer";
            worksheet.Cells["A12"].Value = "Customer S/O #";
            worksheet.Cells["A13"].Value = "PO #";
            worksheet.Cells["A14"].Value = "Date";
            worksheet.Cells["A16"].Value = "Product";
            worksheet.Cells["A17"].Value = "Recipe/Item";
            worksheet.Cells["A18"].Value = "Lot Code";
            worksheet.Cells["A19"].Value = "Manufacture Date";
            worksheet.Cells["A20"].Value = "Manufacturing Site";
            worksheet.Cells["A22"].Value = "% Acid (TA)";
            worksheet.Cells["B22"].Value = AcidMethod;
            worksheet.Cells["B22"].Style.Font.Size = 6;
            worksheet.Cells["A23"].Value = "pH";
            worksheet.Cells["B23"].Value = pHMethod;
            worksheet.Cells["B23"].Style.Font.Size = 6;
            worksheet.Cells["A24"].Value = "Viscosity cps";
            worksheet.Cells["B24"].Value = ViscosityCPSMethod;
            worksheet.Cells["B24"].Style.Font.Size = 6;
            worksheet.Cells["A25"].Value = "Salt %";
            worksheet.Cells["B25"].Value = SaltMethod;
            worksheet.Cells["B25"].Style.Font.Size = 6;
            worksheet.Cells["A27"].Value = "Verified By";
            worksheet.Cells["C27:H27"].Merge = true;
            worksheet.Cells["C27"].Style.HorizontalAlignment = ExcelHorizontalAlignment.Center;

            worksheet.Cells["C29:H33"].Merge = true;
            worksheet.Cells["C29:H33"].Style.HorizontalAlignment = ExcelHorizontalAlignment.Center;
            worksheet.Cells["C29"].Style.WrapText = true;
            worksheet.Cells["C29"].Value = "Please be advised that the results herein are accurate and representative to the best of our knowledge, " +
                "based on current data and information as of the date of this document.  This COA is intended only for the person or company specified " +
                "hereon as the recipient.  This COA shall not be distributed to any person or company other than the intended recipient.  A person or " +
                "company other than the one named hereon shall not rely or make use of the statements and results herein. ";

            for (int i = 1; i <= 6; i++)
            {
                string lotCode = GetLotCode(i + (page - 1) * 6);

                if (string.IsNullOrEmpty(lotCode))
                        continue;

                worksheet.Cells[18, 2 + i].Style.HorizontalAlignment = ExcelHorizontalAlignment.Center;

                string productCode = GetProductCode(lotCode);
                string productName = GetProductName(productCode);

                if (string.IsNullOrEmpty(productName))
                    continue;

                worksheet.Cells[18, 2 + i].Value = lotCode;
                worksheet.Cells[18, 2 + i].Style.Numberformat.Format = "#,##0";

                worksheet.Cells[16, 2 + i].Value = productName;
                worksheet.Cells[16, 2 + i].Style.WrapText = true;
                worksheet.Cells[16, 2 + i].Style.HorizontalAlignment = ExcelHorizontalAlignment.Center;

                string recipeCode = GetRecipeCode(productCode);
                worksheet.Cells[17, 2 + i].Value = recipeCode + "/" + productCode;
                worksheet.Cells[17, 2 + i].Style.HorizontalAlignment = ExcelHorizontalAlignment.Center;

                worksheet.Cells[20, 2 + i].Value = GetManufacturingSite(lotCode);
                worksheet.Cells[20, 2 + i].Style.HorizontalAlignment = ExcelHorizontalAlignment.Center;

                DateTime madeDate = GetMadeDate(productCode, lotCode);
                worksheet.Cells[19, 2 + i].Value = madeDate.ToShortDateString();
                worksheet.Cells[19, 2 + i].Style.HorizontalAlignment = ExcelHorizontalAlignment.Center;

                List<int> titrationIndices = GetTitrationIndices(recipeCode, madeDate, lotCode);
                if (titrationIndices.Count == 0)
                    titrationIndices = GetTitrationIndices(recipeCode, madeDate.AddDays(1), lotCode);

                float acidity = GetTitrationValue(titrationIndices, TitrationOffset.Acidity);
                worksheet.Cells[22, 2 + i].Value = acidity.ToString("0.00");
                if (acidity == -1)
                    worksheet.Cells[22, 2 + i].Style.Font.Color.SetColor(Color.Red);
                worksheet.Cells[22, 2 + i].Style.HorizontalAlignment = ExcelHorizontalAlignment.Center;

                float salt = GetTitrationValue(titrationIndices, TitrationOffset.Salt);
                worksheet.Cells[25, 2 + i].Value = salt.ToString("0.00");
                if (salt == -1)
                    worksheet.Cells[25, 2 + i].Style.Font.Color.SetColor(Color.Red);
                worksheet.Cells[25, 2 + i].Style.HorizontalAlignment = ExcelHorizontalAlignment.Center;

                float ph = GetTitrationValue(titrationIndices, TitrationOffset.pH);
                worksheet.Cells[23, 2 + i].Value = ph.ToString("0.00");
                if (ph == -1)
                    worksheet.Cells[23, 2 + i].Style.Font.Color.SetColor(Color.Red);
                worksheet.Cells[23, 2 + i].Style.HorizontalAlignment = ExcelHorizontalAlignment.Center;

                float viscosity = GetTitrationValue(titrationIndices, TitrationOffset.Viscosity);
                worksheet.Cells[24, 2 + i].Value = viscosity.ToString("0,000");
                if (viscosity == -1)
                    worksheet.Cells[24, 2 + i].Style.Font.Color.SetColor(Color.Red);
                worksheet.Cells[24, 2 + i].Style.HorizontalAlignment = ExcelHorizontalAlignment.Center;

                //Salt T
                // acidity u
                // pH v
                // ReTest_1
            }


        }
        private void PopulateContentsLatitude36(ExcelWorksheet worksheet, int page)
        {
            SaveFile = true;
            CreateTableOfBorders(6, 4, 11, 3, worksheet);
            CreateTableOfBorders(6, 4, 16, 3, worksheet);
            CreateTableOfBorders(6, 8, 21, 3, worksheet);
            CreateTableOfBorders(6, 1, 30, 3, worksheet);

            worksheet.Cells["C11"].Value = "Latitude 36 (Ohio)";
            worksheet.Cells["C12"].Value = Convert.ToInt32(TableauData[1][3]);

            worksheet.Cells["A11"].Value = "Customer";
            worksheet.Cells["A12"].Value = "Customer S/O #";
            worksheet.Cells["A13"].Value = "PO #";
            worksheet.Cells["A14"].Value = "Date";
            worksheet.Cells["A16"].Value = "Product";
            worksheet.Cells["A17"].Value = "Recipe/Item";
            worksheet.Cells["A18"].Value = "Lot Code";
            worksheet.Cells["A19"].Value = "Manufacturing Site";
            worksheet.Cells["A21"].Value = "% Acid (TA)";
            worksheet.Cells["B21"].Value = AcidMethod;
            worksheet.Cells["B21"].Style.Font.Size = 6;
            
            worksheet.Cells["A22"].Value = "pH";
            worksheet.Cells["B22"].Value = pHMethod;
            worksheet.Cells["B22"].Style.Font.Size = 6;
            
            worksheet.Cells["A23"].Value = "Viscosity cps";
            worksheet.Cells["B23"].Value = ViscosityCPSMethod;
            worksheet.Cells["B23"].Style.Font.Size = 6;
            
            worksheet.Cells["A24"].Value = "Yeast cfu/gram";
            worksheet.Cells["B24"].Value = YeastMethod;
            worksheet.Cells["B24"].Style.Font.Size = 6;
            
            worksheet.Cells["A25"].Value = "Mold cfu/gram";
            worksheet.Cells["B25"].Value = MoldMethod;
            worksheet.Cells["B25"].Style.Font.Size = 6;
            
            worksheet.Cells["A26"].Value = "Aerobic cfu/gram";
            worksheet.Cells["B26"].Value = AerobicMethod;
            worksheet.Cells["B26"].Style.Font.Size = 6;
            
            worksheet.Cells["A27"].Value = "Total coliform cfu/gram";
            worksheet.Cells["B27"].Value = ColiformMethod;
            worksheet.Cells["B27"].Style.Font.Size = 6;
           
            worksheet.Cells["A28"].Value = "Lactics cfu/gram";
            worksheet.Cells["B28"].Value = LacticMethod;
            worksheet.Cells["B28"].Style.Font.Size = 6;

            worksheet.Cells["A30"].Value = "Verified By";
            worksheet.Cells["C30:H30"].Merge = true;
            worksheet.Cells["C30"].Style.HorizontalAlignment = ExcelHorizontalAlignment.Center;

            worksheet.Cells["C32:H36"].Merge = true;
            worksheet.Cells["C32"].Style.HorizontalAlignment = ExcelHorizontalAlignment.Center;
            worksheet.Cells["C32"].Style.WrapText = true;
            worksheet.Cells["C32"].Value = "Please be advised that the results herein are accurate and representative to the best of our knowledge, " +
                "based on current data and information as of the date of this document.  This COA is intended only for the person or company specified " +
                "hereon as the recipient.  This COA shall not be distributed to any person or company other than the intended recipient.  A person or " +
                "company other than the one named hereon shall not rely or make use of the statements and results herein. ";

           

            for (int i = 1; i <= 6; i++)
            {
                string lotCode = GetLotCode(i + (page - 1) * 6);

                if (string.IsNullOrEmpty(lotCode))
                        continue;

                worksheet.Cells[16, 2 + i, 30, 2 + i].Style.HorizontalAlignment = ExcelHorizontalAlignment.Center;

                string productCode = GetProductCode(lotCode);
                string productName = GetProductName(productCode);

                if (string.IsNullOrEmpty(productName))
                    continue;

                worksheet.Cells[18, 2 + i].Value = Convert.ToDouble(lotCode);
                worksheet.Cells[18, 2 + i].Style.Numberformat.Format = "0";

                worksheet.Cells[16, 2 + i].Value = productName;
                worksheet.Cells[16, 2 + i].Style.WrapText = true;

                string recipeCode = GetRecipeCode(productCode);
                worksheet.Cells[17, 2 + i].Value = recipeCode + "/" + productCode;

                worksheet.Cells[19, 2 + i].Value = GetManufacturingSite(lotCode);

                DateTime madeDate = GetMadeDate(productCode, lotCode);

                List<int> titrationIndices = GetTitrationIndices(recipeCode, madeDate, lotCode);
                if (titrationIndices.Count == 0)
                    titrationIndices = GetTitrationIndices(recipeCode, madeDate.AddDays(1), lotCode);

                float acidity = GetTitrationValue(titrationIndices, TitrationOffset.Acidity);
                worksheet.Cells[21, 2 + i].Style.Numberformat.Format = "0.00";
                worksheet.Cells[21, 2 + i].Value = acidity;
                if (acidity == -1)
                    worksheet.Cells[21, 2 + i].Style.Font.Color.SetColor(Color.Red);


                float ph = GetTitrationValue(titrationIndices, TitrationOffset.pH);
                worksheet.Cells[22, 2 + i].Value = ph;
                worksheet.Cells[22, 2 + i].Style.Numberformat.Format = "0.00";
                if (ph == -1)
                    worksheet.Cells[22, 2 + i].Style.Font.Color.SetColor(Color.Red);

                float viscosity = GetTitrationValue(titrationIndices, TitrationOffset.Viscosity);
                worksheet.Cells[23, 2 + i].Value = viscosity;
                worksheet.Cells[23, 2 + i].Style.Numberformat.Format = "0,000";
                if (viscosity == -1)
                    worksheet.Cells[23, 2 + i].Style.Font.Color.SetColor(Color.Red);

                List<int> microIndices = GetMicroIndices(recipeCode, lotCode, productCode, madeDate);

                //worksheet.Cells[29, 1].Value = "Micro Index";
                //worksheet.Cells[31, 1].Value = "Titration Index";
                //foreach (int index in microIndices)
                //{
                //    worksheet.Cells[29, 2 + i].Value += Convert.ToString(index + 1) + ", ";
                //}
                //foreach (int index in titrationIndices)
                //{
                //    worksheet.Cells[31, 2 + i].Value += Convert.ToString(index + 1) + ", ";
                //}

                if (microIndices.Count == 0)
                {
                    worksheet.Cells[24, 2 + i, 28, 2 + i].Value = "No result";
                    worksheet.Cells[24, 2 + i, 28, 2 + i].Style.Font.Color.SetColor(Color.Red);
                }
                else
                {
                    string factoryCode = GetFactoryCode(lotCode);

                    int yeastCount = GetMicroValue(microIndices, MicroOffset.Yeast);

                    if(MicroValueInSpec(yeastCount, MicroOffset.Yeast) == false)
                        worksheet.Cells[24, 2 + i].Style.Font.Color.SetColor(Color.Yellow);

                    if (yeastCount == 0)
                        worksheet.Cells[24, 2 + i].Value = "<10";
                    else if (yeastCount > 0)
                        worksheet.Cells[24, 2 + i].Value = yeastCount;
                    else if (yeastCount == -1)
                        worksheet.Cells[24, 2 + i].Value = "N/A";

                    worksheet.Cells[24, 2 + i].Style.Numberformat.Format = "#,##0";


                    int moldCount = GetMicroValue(microIndices, MicroOffset.Mold);

                    if (MicroValueInSpec(moldCount, MicroOffset.Mold) == false)
                        worksheet.Cells[25, 2 + i].Style.Font.Color.SetColor(Color.Yellow);
                    
                    if (moldCount == 0)
                        worksheet.Cells[25, 2 + i].Value = "<10";
                    else if (moldCount > 0)
                        worksheet.Cells[25, 2 + i].Value = moldCount;
                    else if (moldCount == -1)
                        worksheet.Cells[25, 2 + i].Value = "N/A";

                    worksheet.Cells[25, 2 + i].Style.Numberformat.Format = "#,##0";

                    int aerobicCount = GetMicroValue(microIndices, MicroOffset.Aerobic);

                    if (MicroValueInSpec(aerobicCount, MicroOffset.Aerobic) == false)
                        worksheet.Cells[26, 2 + i].Style.Font.Color.SetColor(Color.Yellow);

                    if (aerobicCount == 0 && GetFactoryCode(lotCode) == "H1")
                        worksheet.Cells[26, 2 + i].Value = "<10";
                    else if (aerobicCount == 0)
                        worksheet.Cells[26, 2 + i].Value = "<100";
                    else if (aerobicCount > 0)
                        worksheet.Cells[26, 2 + i].Value = aerobicCount;
                    else if (aerobicCount == -1)
                        worksheet.Cells[26, 2 + i].Value = "N/A";

                    worksheet.Cells[26, 2 + i].Style.Numberformat.Format = "#,##0";

                    int coliformCount = GetMicroValue(microIndices, MicroOffset.Coliform);

                    if (MicroValueInSpec(coliformCount, MicroOffset.Coliform) == false)
                        worksheet.Cells[27, 2 + i].Style.Font.Color.SetColor(Color.Yellow);

                    if (coliformCount == 0)
                        worksheet.Cells[27, 2 + i].Value = "<10";
                    else if (coliformCount > 0)
                        worksheet.Cells[27, 2 + i].Value = coliformCount;
                    else if (coliformCount == -1)
                        worksheet.Cells[27, 2 + i].Value = "N/A";

                    worksheet.Cells[27, 2 + i].Style.Numberformat.Format = "#,##0";

                    int lacticCount = GetMicroValue(microIndices, MicroOffset.Lactic);

                    if (MicroValueInSpec(lacticCount, MicroOffset.Lactic) == false)
                        worksheet.Cells[28, 2 + i].Style.Font.Color.SetColor(Color.Yellow);

                    if (lacticCount == 0)
                        worksheet.Cells[28, 2 + i].Value = "<10";
                    else if (lacticCount > 0)
                        worksheet.Cells[28, 2 + i].Value = lacticCount;
                    else if (lacticCount == -1)
                        worksheet.Cells[28, 2 + i].Value = "N/A";

                    worksheet.Cells[28, 2 + i].Style.Numberformat.Format = "#,##0";
                }

            }
        }
        private void PopulateContentsKootenaiAndCheese(ExcelWorksheet worksheet, int page)
        {
            CreateTableOfBorders(6, 3, 11, 3, worksheet);
            CreateTableOfBorders(6, 3, 15, 3, worksheet);
            CreateTableOfBorders(6, 6, 19, 3, worksheet);
            CreateTableOfBorders(6, 1, 26, 3, worksheet);

            worksheet.Cells["C11:H26"].Style.HorizontalAlignment = ExcelHorizontalAlignment.Center;
            worksheet.Cells["C11:H26"].Style.VerticalAlignment = ExcelVerticalAlignment.Center;

            worksheet.Cells["A11"].Value = "Item";
            worksheet.Cells["A12"].Value = "Made Date";
            worksheet.Cells["A13"].Value = "Date";
            worksheet.Cells["A15"].Value = "Best by";
            worksheet.Cells["A16"].Value = "Batch";
            worksheet.Cells["A17"].Value = "Lot";

            worksheet.Cells["C11"].Value = InternalCOAData[1];
            worksheet.Cells["C12"].Value = Convert.ToDateTime(InternalCOAData[0]).ToShortDateString();
            worksheet.Cells["C13"].Value = DateTime.Now.ToShortDateString();
            worksheet.Cells["C14"].Value = string.Empty;
            worksheet.Cells["C14:H14"].Merge = false;
            worksheet.Cells["C14:H14"].Style.Font.SetFromFont(new Font("Calibri", 11, FontStyle.Regular));
            worksheet.Cells["C14:H14"].Style.HorizontalAlignment = ExcelHorizontalAlignment.Center;

            worksheet.Cells["A19"].Value = "Yeast cfu/gram";
            worksheet.Cells["B19"].Value = YeastMethod;

            worksheet.Cells["A20"].Value = "Mold cfu/gram";
            worksheet.Cells["B20"].Value = MoldMethod;

            worksheet.Cells["A21"].Value = "Aerobic cfu/gram";
            worksheet.Cells["B21"].Value = AerobicMethod;
            
            worksheet.Cells["A22"].Value = "Total coliform cfu/gram";
            worksheet.Cells["B22"].Value = ColiformMethod;

            worksheet.Cells["A23"].Value = "E. Coliform cfu/gram";
            worksheet.Cells["B23"].Value = EColiMethod;

            worksheet.Cells["A24"].Value = "Lactics cfu/gram";
            worksheet.Cells["B24"].Value = LacticMethod;
            
            worksheet.Cells["A26"].Value = "Verified By";
            worksheet.Cells["C26:H26"].Merge = true;
            worksheet.Cells["C26"].Style.HorizontalAlignment = ExcelHorizontalAlignment.Center;


            List<int> microIndices = GetMicroIndices(InternalCOAData[1], Convert.ToDateTime(InternalCOAData[0]), "K1");

            for (int i = 1; i <= 6; i++)
            {
                if(i > microIndices.Count)
                {
                    continue;
                }

                if(MicroResults[microIndices[i - 1]][11] == "*")
                {
                    worksheet.Cells[16, 2 + i].Value = "N/A";
                }
                else
                {
                    worksheet.Cells[16, 2 + i].Value = MicroResults[microIndices[i - 1]][11];
                }

                if (MicroResults[microIndices[i - 1]][12] == "*")
                {
                    worksheet.Cells[15, 2 + i].Value = "N/A";
                }
                else
                {
                    worksheet.Cells[15, 2 + i].Value = MicroResults[microIndices[i - 1]][12];
                }

                if (MicroResults[microIndices[i - 1]][17] == "*")
                {
                    worksheet.Cells[17, 2 + i].Value = "N/A";
                }
                else
                {
                    worksheet.Cells[17, 2 + i].Value = MicroResults[microIndices[i - 1]][17];
                }

                if (microIndices.Count == 0)
                {
                    worksheet.Cells[19, 2 + i, 24, 2 + i].Value = "No result";
                    worksheet.Cells[19, 2 + i, 24, 2 + i].Style.Font.Color.SetColor(Color.Red);
                }
                else
                {
                    int yeastCount = GetMicroValue(microIndices, MicroOffset.Yeast);

                    if (MicroValueInSpec(yeastCount, MicroOffset.Yeast) == false)
                        worksheet.Cells[19, 2 + i].Style.Font.Color.SetColor(Color.Yellow);

                    if (yeastCount == 0)
                        worksheet.Cells[19, 2 + i].Value = "<10";
                    else if (yeastCount > 0)
                        worksheet.Cells[19, 2 + i].Value = yeastCount;
                    else if (yeastCount == -1)
                        worksheet.Cells[19, 2 + i].Value = "N/A";

                    worksheet.Cells[19, 2 + i].Style.Numberformat.Format = "#,##0";


                    int moldCount = GetMicroValue(microIndices, MicroOffset.Mold);

                    if (MicroValueInSpec(moldCount, MicroOffset.Mold) == false)
                        worksheet.Cells[20, 2 + i].Style.Font.Color.SetColor(Color.Yellow);

                    if (moldCount == 0)
                        worksheet.Cells[20, 2 + i].Value = "<10";
                    else if (moldCount > 0)
                        worksheet.Cells[20, 2 + i].Value = moldCount;
                    else if (moldCount == -1)
                        worksheet.Cells[20, 2 + i].Value = "N/A";

                    worksheet.Cells[20, 2 + i].Style.Numberformat.Format = "#,##0";

                    int aerobicCount = GetMicroValue(microIndices, MicroOffset.Aerobic);

                    if (MicroValueInSpec(aerobicCount, MicroOffset.Aerobic) == false)
                        worksheet.Cells[21, 2 + i].Style.Font.Color.SetColor(Color.Yellow);

                    if (aerobicCount == 0)
                        worksheet.Cells[21, 2 + i].Value = "<100";
                    else if (aerobicCount > 0)
                        worksheet.Cells[21, 2 + i].Value = aerobicCount;
                    else if (aerobicCount == -1)
                        worksheet.Cells[21, 2 + i].Value = "N/A";

                    worksheet.Cells[21, 2 + i].Style.Numberformat.Format = "#,##0";

                    int coliformCount = GetMicroValue(microIndices, MicroOffset.Coliform);

                    if (MicroValueInSpec(coliformCount, MicroOffset.Coliform) == false)
                        worksheet.Cells[22, 2 + i].Style.Font.Color.SetColor(Color.Yellow);

                    if (coliformCount == 0)
                        worksheet.Cells[22, 2 + i].Value = "<10";
                    else if (coliformCount > 0)
                        worksheet.Cells[22, 2 + i].Value = coliformCount;
                    else if (coliformCount == -1)
                        worksheet.Cells[22, 2 + i].Value = "N/A";

                    worksheet.Cells[22, 2 + i].Style.Numberformat.Format = "#,##0";

                    int eColiformCount = GetMicroValue(microIndices, MicroOffset.EColiform);

                    if (MicroValueInSpec(eColiformCount, MicroOffset.EColiform) == false)
                        worksheet.Cells[23, 2 + i].Style.Font.Color.SetColor(Color.Yellow);

                    if (coliformCount == 0)
                        worksheet.Cells[23, 2 + i].Value = "<10";
                    else if (coliformCount > 0)
                        worksheet.Cells[23, 2 + i].Value = coliformCount;
                    else if (coliformCount == -1)
                        worksheet.Cells[23, 2 + i].Value = "N/A";

                    worksheet.Cells[23, 2 + i].Style.Numberformat.Format = "#,##0";

                    int lacticCount = GetMicroValue(microIndices, MicroOffset.Lactic);

                    if (MicroValueInSpec(lacticCount, MicroOffset.Lactic) == false)
                        worksheet.Cells[24, 2 + i].Style.Font.Color.SetColor(Color.Yellow);

                    if (lacticCount == 0)
                        worksheet.Cells[24, 2 + i].Value = "<10";
                    else if (lacticCount > 0)
                        worksheet.Cells[24, 2 + i].Value = lacticCount;
                    else if (lacticCount == -1)
                        worksheet.Cells[24, 2 + i].Value = "N/A";

                    worksheet.Cells[24, 2 + i].Style.Numberformat.Format = "#,##0";

                    if(yeastCount != -1 || moldCount != -1 || aerobicCount != -1 || coliformCount != -1 || eColiformCount != -1 || lacticCount != -1 )
                        SaveFile = true;
                }
            }
        }
        /// <summary>
        /// Creates borders around each cell in the specified range
        /// </summary>
        /// <param name="width"></param>
        /// <param name="height"></param>
        /// <param name="rowStart">Begins at 1</param>
        /// <param name="startColumn">Begins at 1</param>
        /// <param name="worksheet"></param>
        private void CreateTableOfBorders(int width, int height, int rowStart, int startColumn, ExcelWorksheet worksheet)
        {
            for (int i = 0; i < height; i++)
            {
                for (int j = 0; j < width; j++)
                {
                    worksheet.Cells[rowStart + i, startColumn + j].Style.Border.BorderAround(ExcelBorderStyle.Thin);
                }
            }
        }
        /// <summary>
        /// 
        /// </summary>
        /// <param name="tableauItemRow">Is zero-based</param>
        private string GetProductCode(string lotCode)
        {
            string productCode = lotCode[0].ToString();
            productCode += lotCode[1];
            productCode += lotCode[2];
            productCode += lotCode[3];
            productCode += lotCode[4];

            return productCode;
        }
        private string GetProductName(string productCode)
        {
            foreach (List<string> line in FinishedGoods)
            {
                if (line[0] == productCode)
                {
                    return line[1];
                }
            }
            return string.Empty;
        }
        private string GetRecipeCode(string productCode)
        {
            foreach (List<string> line in FinishedGoods)
            {
                if (line[0] == productCode)
                {
                    return line[3];
                }
            }
            return "Could not locate";
        }
        private string GetManufacturingSite(string lotCode)
        {
            string locationCode = lotCode[5].ToString();
            locationCode += lotCode[6];

            if (locationCode == "01")
            {
                return "Sandpoint, ID";
            }
            else if (locationCode == "02")
            {
                return "Lowell, MI";
            }
            else if (locationCode == "03")
            {
                return "Hurricane, UT";
            }
            else
            {
                return locationCode;
            }

        }
        /// <summary>
        /// Retrieves lot code from WorkbookData's tableau data by given index
        /// </summary>
        /// <param name="tableauItemRow">The 0-based index to retrieve from</param>
        /// <returns></returns>
        private string GetLotCode(int tableauItemRow)
        {
            if (tableauItemRow > TableauData.Count - 1)
            {
                return string.Empty;
            }
            else if (string.IsNullOrEmpty(TableauData[tableauItemRow][0]))
            {
                return string.Empty;
            }

            if (TableauData[tableauItemRow][0].Contains(' '))
            {
                string trimmedLotCode = string.Empty;

                foreach (char character in TableauData[tableauItemRow][0])
                {
                    if (character != ' ')
                    {
                        trimmedLotCode += character.ToString();
                    }
                }

                return trimmedLotCode;
            }

            return TableauData[tableauItemRow][0];
        }
        private DateTime GetMadeDate(string productCode, string lotCode)
        {
            int daysToExpiry = 0;

            foreach (List<string> line in FinishedGoods)
            {
                if (line[0] == productCode)
                {
                    daysToExpiry = Convert.ToInt32(line[2]);
                }
            }

            int expiryDay = Convert.ToInt32(lotCode[9].ToString()) * 10;
            expiryDay += Convert.ToInt32(lotCode[10].ToString());

            int expiryMonth = Convert.ToInt32(lotCode[7].ToString()) * 10;
            expiryMonth += Convert.ToInt32(lotCode[8].ToString());

            int expiryYear = Convert.ToInt32(lotCode[11].ToString()) * 10;
            expiryYear += Convert.ToInt32(lotCode[12].ToString());
            expiryYear += 2000;

            return new DateTime(expiryYear, expiryMonth, expiryDay).AddDays(daysToExpiry * -1);
        }
        private List<int> GetTitrationIndices(string recipeCode, DateTime madeDate, string lotCode)
        {
            List<int> indices = new List<int>();
            string jobNumber = string.Empty;
            string factoryCode = GetFactoryCode(lotCode);
            string madeDateAsString = madeDate.ToShortDateString();
            string madeDateAsTwoDigitYearString = madeDate.ToString("M/d/yy");

            for (int i = 0; i < TitrationResults.Count; i++)
            {
                if (TitrationResults[i][2] == factoryCode && (TitrationResults[i][0] == madeDateAsString ||
                    (TitrationResults[i][0] == madeDateAsTwoDigitYearString)) && TitrationResults[i][4] == recipeCode)
                {
                    indices.Add(i);

                    if (string.IsNullOrEmpty(jobNumber))
                    {
                        jobNumber = TitrationResults[i][3];
                    }
                }
            }

            // Runs a second check under a different factory
            if (indices.Count == 0)
            {
                if (factoryCode == "H1")
                    factoryCode = "L1";
                else if (factoryCode == "L1")
                    factoryCode = "H1";

                for (int i = 0; i < TitrationResults.Count; i++)
                {
                    if (TitrationResults[i][2] == factoryCode && (TitrationResults[i][0] == madeDateAsString ||
                        (TitrationResults[i][0] == madeDateAsTwoDigitYearString)) && TitrationResults[i][4] == recipeCode)
                    {
                        indices.Add(i);

                        if (string.IsNullOrEmpty(jobNumber))
                        {
                            jobNumber = TitrationResults[i][3];
                        }
                    }
                }
            }


            return indices;
        }
        private float GetTitrationValue(List<int> indices, TitrationOffset offset)
        {
            List<float> results = new List<float>();

            foreach (int index in indices)
            {
                for (int i = 0; i < TitrationResults[index].Count; i++)
                {
                    string value = TitrationResults[index][i];

                    if (value == "Original" || value == "ReTest_1" || value == "ReTest_2" || value == "ReTest_3" || value == "ReTest_4" || value == "ReTest_5")
                    {
                        if (TitrationResults[index][i + (int)offset] != "*")
                            results.Add(Convert.ToSingle(TitrationResults[index][i + (int)offset]));
                    }
                }
            }

            float sum = 0;

            foreach (float value in results)
            {
                sum += value;
            }

            float average = sum / results.Count;

            int closestValueIndex = 0;
            float smallestDifference = 10000000000;
            float currentDifference = 0;

            for (int i = 0; i < results.Count; i++)
            {
                currentDifference = results[i] - average >= 0 ? results[i] - average : average - results[i];

                if (currentDifference < smallestDifference)
                {
                    smallestDifference = currentDifference;
                    closestValueIndex = i;
                }

            }

            return results.Count == 0 ? -1 : results[closestValueIndex];
        }
        private List<int> GetMicroIndices(string recipeCode, string lotCode, string productCode, DateTime madeDate)
        {
            List<int> indices = new List<int>();

            string factoryCode = GetFactoryCode(lotCode);
            string madeDateAsString = madeDate.ToShortDateString();
            string madeDateAsTwoDigitYearString = madeDate.ToString("M/d/yy");

            for (int i = 0; i < MicroResults.Count; i++)
            {
                if (MicroResults[i][0] == factoryCode && MicroResults[i][7] == recipeCode && MicroResults[i][10] == productCode &&
                    (madeDateAsString == MicroResults[i][9] || madeDateAsTwoDigitYearString == MicroResults[i][9]))
                {
                    indices.Add(i);
                }
            }
            return indices;
        }
        private List<int> GetMicroIndices(string productCode, DateTime madeDate, string supplier)
        {
            List<int> indices = new List<int>();

            string madeDateAsString = madeDate.ToShortDateString();
            string madeDateAsTwoDigitYearString = madeDate.ToString("M/d/yy");

            for (int i = 0; i < MicroResults.Count; i++)
            {
                if (MicroResults[i][16] == supplier && MicroResults[i][10] == productCode &&
                    (madeDateAsString == MicroResults[i][9] || madeDateAsTwoDigitYearString == MicroResults[i][9]))
                {
                    indices.Add(i);
                }
            }
            return indices;
        }

        private int GetMicroValue(List<int> indices, MicroOffset offset)
        {
            List<int> microValues = new List<int>();

            foreach (int index in indices)
            {
                for (int i = 0; i < MicroResults[index].Count; i++)
                {
                    if (MicroResults[index][i] == "HURRICANE" || MicroResults[index][i] == "Hurricane" || MicroResults[index][i] == "Lowell" || MicroResults[index][i] == "Sandpoint")
                    {
                        if (string.IsNullOrEmpty(MicroResults[index][i + (int)offset]) || MicroResults[index][i + (int)offset] == "*")
                            continue;
                        else

                            microValues.Add(Convert.ToInt32(MicroResults[index][i + (int)offset].Trim()));
                    }
                }
            }

            int largestValue = -1;

            foreach (int value in FilterMicroValues(microValues, offset))
            {
                if (value > largestValue)
                {
                    largestValue = value;
                }
            }

            return largestValue;
        }
        /// <summary>
        /// Filters out-of-spec values provided that at least one value is in-spec, otherwise the provided list is returned
        /// </summary>
        /// <param name="unsortedValues"></param>
        /// <param name="offset"></param>
        /// <returns></returns>
        private List<int> FilterMicroValues(List<int> unsortedValues, MicroOffset offset)
        {
            List<int> sortedValues = new List<int>();

                foreach(int value in unsortedValues)
                {
                if (offset == MicroOffset.Aerobic && value < 100000) // 100k
                    sortedValues.Add(value);
                else if (offset == MicroOffset.Coliform && value < 100)
                    sortedValues.Add(value);
                else if (offset == MicroOffset.Lactic && value < 1000)
                    sortedValues.Add(value);
                else if (offset == MicroOffset.Mold && value < 1000)
                    sortedValues.Add(value);
                else if (offset == MicroOffset.Yeast && value < 1000)
                    sortedValues.Add(value);
                else if (offset == MicroOffset.EColiform && value == 0)
                    sortedValues.Add(value);
                }

                if (sortedValues.Count != 0)
                {
                    return sortedValues;
                }
                else
                {
                    return unsortedValues;
                }
            
        }
        private bool MicroValueInSpec(int value, MicroOffset offset)
        {
            if(offset == MicroOffset.Aerobic)
            {
                if(value < 100000)
                    return true;
                else
                    return false;
            }
            else if(offset == MicroOffset.Coliform)
            {
                if(value < 100)
                    return true;
                else
                    return false;
            }
            else if(offset == MicroOffset.Lactic)
            {
                if(value < 1000)
                    return true;
                else
                    return false;
            }
            else if (offset == MicroOffset.Mold)
            {
                if(value < 1000)
                    return true;
                else
                    return false;
            }
            else // when (offset == MicroOffset.Yeast)
            {
                if(value < 1000)
                    return true;
                else
                    return false;
            }
        }
        /// <summary>
        /// Retrieves shelf life for a given recipe from the Recipes list
        /// </summary>
        /// <param name="recipeCode"></param>
        /// <returns></returns>
        private string GetFactoryCode(string lotCode)
        {
            string locationCode = lotCode[5].ToString();
            locationCode += lotCode[6];

            if (locationCode == "01")
                return "S1";
            else if (locationCode == "02")
                return "L1";
            else if (locationCode == "03")
                return "H1";
            else
                return string.Empty;
        }
        private bool DoesFilterInvalidateDocument()
        {
            if(WorkbookTemplate.CustomFilters.Count > 0)
            {
                if(WorkbookTemplate.SelectedAlgorithm == Templates.Template.Algorithm.Standard) // Get tableau lots if true
                {
                    List<string> lotCodes = new List<string>();

                    for (int i = 1; i < TableauData.Count; i++)
                    {
                        lotCodes.Add(GetLotCode(i));
                    }
                }
                foreach(Templates.CustomFilter filter in WorkbookTemplate.CustomFilters)
                {
                    if (filter.IsValidFilter)
                    {
                        switch(filter.ContentItem)
                        {
                            case Templates.Template.ContentItems.RecipeAndItem:
                                if(filter.FilterType == Templates.CustomFilter.FilterTypes.Whitelist)
                                {
                                    if(WorkbookTemplate.SelectedAlgorithm == Templates.Template.Algorithm.Standard)
                                    {

                                    }
                                    else // SelectedAlgorithm == Algorithm.FromDateOnwards
                                    {

                                    }
                                }
                                else // FilterType == Blacklist
                                {

                                }
                                break;
                            default:
                                break;
                        }
                    }
                }
            }

            return false;
        }
    }
}
