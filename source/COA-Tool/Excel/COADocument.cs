using System;
using System.Collections.Generic;
using System.Text;
using OfficeOpenXml;
using System.Xml;
using System.Drawing;
using OfficeOpenXml.Style;
using System.IO;
using OfficeOpenXml.Drawing;
using System.Linq;


namespace CoA_Tool.Excel
{
    /// <summary>
    /// Responsible for generating final output document
    /// </summary>
    class COADocument
    {
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

        // DateTimes
        public DateTime StartDate;

        // Hashsets
        private HashSet<string> BatchIndicesToIgnore;

        //  Objects
        private Templates.Template WorkbookTemplate;
        private CSV.SalesOrder SalesOrder { get; set; }
        private CSV.NWAData NWAData { get; set; }
        private CSV.TableauData TableauData { get; set; }
        private FinishedGoodsData FinishedGoodsData { get; set; }

        // Constructor
        public COADocument(Templates.Template template, CSV.SalesOrder salesOrder, CSV.NWAData nwaData, CSV.TableauData tableau, FinishedGoodsData finishedGoods)
        {
            WorkbookTemplate = template;
            SalesOrder = salesOrder;
            NWAData = nwaData;
            TableauData = tableau;
            FinishedGoodsData = finishedGoods;
        }
        /// <summary>
        /// Generate a workbook using the standard algorithm
        /// </summary>
        public void StandardGeneration()
        {
            using (ExcelPackage package = new ExcelPackage())
            {
                int pageCount = 0;

                pageCount = (SalesOrder.Lots.Count- 1) / 6;
                if ((SalesOrder.Lots.Count - 1) % 6 > 0)
                {
                    pageCount++;
                }
                
                for (int currentPage = 1; currentPage <= pageCount; currentPage++)
                {
                    ExcelWorksheet worksheet = package.Workbook.Worksheets.Add("Page " + currentPage);
                    PopulateGeneralWorksheetData(worksheet);

                    List<string> lotsToProcess = new List<string>();

                    for (int i = 6 * (currentPage - 1); i < 6 * currentPage; i++)
                    {
                        lotsToProcess.Add(SalesOrder.Lots[i]); // i is zero-based
                    }

                    List<List<int>> TitrationIndices = new List<List<int>>();

                    foreach(string lot in lotsToProcess)
                    {
                        string recipeCode = FinishedGoodsData.RecipeCodeFor(CSV.SalesOrder.ProductCodeFromLot(lot));
                        DateTime madeDate = FinishedGoodsData.GetMadeDate(lot);
                        string factoryCode = SalesOrder.ManufacturingSiteFromLot(lot);
                        TitrationIndices.Add(NWAData.FindTitrationIndices(recipeCode, madeDate, factoryCode));
                    }
                    // TODO: Fetch micro and titration indices and pass them to method for greater flexibility
                    PopulateMainWorksheetContents(worksheet, currentPage);
                    
                }
                
                if (Directory.Exists(Environment.GetFolderPath(Environment.SpecialFolder.Desktop) + "/CoAs") == false)
                {
                    Directory.CreateDirectory(Environment.GetFolderPath(Environment.SpecialFolder.Desktop) + "/CoAs");
                }
                        
                package.SaveAs(new FileInfo(Environment.GetFolderPath(Environment.SpecialFolder.Desktop) + "/CoAs/" + SalesOrder.OrderNumber + ".xlsx"));
            }
        }
        
        /// /// <summary>
        /// Sets static data and settings for the worksheet
        /// </summary>
        /// <param name="targetWorksheet"></param>
        private void PopulateGeneralWorksheetData(ExcelWorksheet targetWorksheet)
        {
            targetWorksheet.Cells.Style.Font.SetFromFont(new Font("Calibri", 11));

            targetWorksheet.View.ShowGridLines = false;

            targetWorksheet.Cells["A1:H55"].Style.HorizontalAlignment = ExcelHorizontalAlignment.Center;
            targetWorksheet.Cells["A1:H55"].Style.VerticalAlignment = ExcelVerticalAlignment.Center;

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
            logo.SetSize(234, 124); // Is affected by row and column resizing
            logo.SetPosition(0, 301); // Is affected by row and column resizing

            targetWorksheet.Cells[11, 1, 60, 2].Style.Font.Size = 9;
            targetWorksheet.Cells[11, 1, 60, 8].Style.VerticalAlignment = ExcelVerticalAlignment.Center;

            targetWorksheet.PrinterSettings.FitToPage = true;

            targetWorksheet.HeaderFooter.EvenFooter.LeftAlignedText = "created: 04/22/2016";
            targetWorksheet.HeaderFooter.EvenFooter.RightAlignedText = "10/31/2019 REV 03 F142-087";
            targetWorksheet.HeaderFooter.OddFooter.LeftAlignedText = "created: 04/22/2016";
            targetWorksheet.HeaderFooter.OddFooter.RightAlignedText = "10/31/2019 REV 03 F142-087";

        }
        /// <summary>
        /// Populates dynamic content for the worksheet
        /// </summary>
        /// <param name="targetWorksheet"></param>
        private void PopulateMainWorksheetContents(ExcelWorksheet targetWorksheet, int currentPage)
        {
            targetWorksheet.Cells["A8"].Value = "Certificate of Analysis";
            targetWorksheet.Cells["A8"].Style.Font.SetFromFont(new Font("Calibri", 26, FontStyle.Bold));
            targetWorksheet.Cells["A8"].Style.HorizontalAlignment = ExcelHorizontalAlignment.Center;
            targetWorksheet.Cells["A8:H9"].Merge = true;

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
                targetWorksheet.Cells[currentRow, 3].Value = SalesOrder.OrderNumber;
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
                lotsToProcess.Add(SalesOrder.Lots[i]); // i is zero-based
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
