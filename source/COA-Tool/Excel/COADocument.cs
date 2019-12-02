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
using CoA_Tool.Utility;

namespace CoA_Tool.Excel
{
    /// <summary>
    /// Responsible for generating the final output document
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

                bool containsInvalidLot = false;
                string invalidLotValue = "";

                for (int currentPage = 1; currentPage <= pageCount; currentPage++) 
                {
                    ExcelWorksheet worksheet = package.Workbook.Worksheets.Add("Page " + currentPage);
                    PopulateGeneralWorksheetData(worksheet);

                    List<string> lotsToProcess = new List<string>();

                    for (int i = 6 * (currentPage - 1); i < 6 * currentPage; i++)
                    {
                        if (i + 1 > SalesOrder.Lots.Count) // TODO: Review this
                        {
                            break;
                        }
                        else
                        {
                            lotsToProcess.Add(SalesOrder.Lots[i]); // i is zero-based
                        }
                    }
                    
                    List<List<int>> titrationIndices = new List<List<int>>();  // Find indices needed for each worksheet
                    List<List<int>> microIndices = new List<List<int>>(); // // Find indices needed for each worksheet


                    foreach (string lotCode in lotsToProcess)
                    {
                        if(FinishedGoodsData.RecipeCodeExists(Lot.ProductCode(lotCode), out string recipeCode) == false)
                        {
                            containsInvalidLot = true;
                            invalidLotValue = lotCode;
                        }
                        
                        DateTime madeDate = FinishedGoodsData.GetMadeDate(lotCode);
                        string factoryCode = Lot.FactoryCode(lotCode);

                        titrationIndices.Add(NWAData.TitrationIndices(recipeCode, madeDate, factoryCode));
                        microIndices.Add(NWAData.MicroIndices(recipeCode, lotCode, madeDate));
                    }

                    if(containsInvalidLot == false)
                    {
                        PopulateMainWorksheetContents(worksheet, titrationIndices, microIndices);
                    }
                }
                
                if (Directory.Exists(Environment.GetFolderPath(Environment.SpecialFolder.Desktop) + "/CoAs") == false)
                {
                    Directory.CreateDirectory(Environment.GetFolderPath(Environment.SpecialFolder.Desktop) + "/CoAs");
                }

                string fileName;

                if (containsInvalidLot)
                {
                    fileName = SalesOrder.OrderNumber + " (Lot " + invalidLotValue + " invalid)" + ".xlsx";
                }
                else
                {
                    fileName = SalesOrder.OrderNumber + ".xlsx";
                }

                bool fileSaved = false;

                List<string> options = new List<string>();
                options.Add("File has been closed");

                do // TODO: Replace with asynchronous solution when async generation is developed to increase resiliency
                {
                    // *Typically* only one document, if any, is opened by the user.  
                    // Meaning that this should be able to hold things over until a more elegant solution is implemented.
                    try
                    {
                        package.SaveAs(new FileInfo(Environment.GetFolderPath(Environment.SpecialFolder.Desktop) + "/CoAs/" + fileName));

                        if (File.Exists(Environment.GetFolderPath(Environment.SpecialFolder.Desktop) + "/CoAs/" + fileName))
                        {
                            fileSaved = true;
                        }
                    }
                    catch (IOException)
                    {
                        new ConsoleInteraction.SelectionMenu(options, " Select:",
                            "\"" + fileName + "\" is being accessed.  Please close the file before continuing.");
                    }
                    catch(InvalidOperationException)
                    {
                        new ConsoleInteraction.SelectionMenu(options, " Select:",
                            "\"" + fileName + "\" is being accessed.  Please close the file before continuing.");
                    }
                } while (fileSaved == false);

                TableauData.MoveBetweenDirectories(TableauData.DesktopSubDirectoryPath + "\\1) New Batch\\" + SalesOrder.OrderNumber + ".csv", CSV.TableauData.LotDirectory.PreviousBatch);
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
            targetWorksheet.Cells["A1:H55"].Style.WrapText = true;

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
        /// <param name="targetWorksheet">The worksheet to populate</param>
        /// <param name="currentPage"></param>
        /// <param name="titrationIndices"></param>
        /// <param name="microIndices"></param>
        private void PopulateMainWorksheetContents(ExcelWorksheet targetWorksheet, List<List<int>> titrationIndices, List<List<int>> microIndices)
        {
            int worksheetNumber = Convert.ToInt32(targetWorksheet.Name.Substring(5)); // Extracts worksheet number from the worksheet name; formatted as "Page n"

            int itemsInWorksheet = titrationIndices.Count;

            if (itemsInWorksheet == 0)
            {
                itemsInWorksheet = microIndices.Count;
            }

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

            if (WorkbookTemplate.IncludeSalesOrder)
            {
                sizeOfFirstContentBlock++;

                targetWorksheet.Cells[currentRow, 1, currentRow, 2].Merge = true;
                targetWorksheet.Cells[currentRow, 1].Value = "Customer S/O #";
                targetWorksheet.Cells[currentRow, 3, currentRow, 8].Merge = true;
                if (Int64.TryParse(SalesOrder.OrderNumber, out long salesOrderAsLong) == true)
                {
                    targetWorksheet.Cells[currentRow, 3].Value = salesOrderAsLong;
                }
                else
                {
                    targetWorksheet.Cells[currentRow, 3].Value = SalesOrder.OrderNumber;
                }

                currentRow++;
            }

            if (WorkbookTemplate.IncludePurchaseOrder)
            {
                sizeOfFirstContentBlock++;

                targetWorksheet.Cells[currentRow, 1, currentRow, 2].Merge = true;
                targetWorksheet.Cells[currentRow, 1].Value = "PO #";
                targetWorksheet.Cells[currentRow, 3, currentRow, 8].Merge = true;
                targetWorksheet.Cells[currentRow, 3].Style.Numberformat.Format = "0";

                currentRow++;
            }

            if (WorkbookTemplate.IncludeGenerationDate)
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
            if (sizeOfFirstContentBlock > 0)
            {
                targetWorksheet.Cells[10, 3, 10, 8].Style.Border.Bottom.Style = ExcelBorderStyle.Medium;
                targetWorksheet.Cells[11, 3, 10 + sizeOfFirstContentBlock, 8].Style.Font.Size = 12;
                targetWorksheet.Cells[11, 3, 10 + sizeOfFirstContentBlock, 8].Style.Font.Bold = true;
                targetWorksheet.Cells[11, 3, 10 + sizeOfFirstContentBlock, 8].Style.HorizontalAlignment = ExcelHorizontalAlignment.Center;

                CreateTableOfBorders(6, sizeOfFirstContentBlock, 11, 3, targetWorksheet);
            }

            // For second content block

            if (sizeOfFirstContentBlock > 0)
            {
                currentRow += 2; // An empty space between blocks and an empty space for the block's header if the block size > 0
            }

            else
            {
                currentRow = 12; // Allows the second block's header to use row 11
            }

            int sizeOfSecondContentBlock = 0;

            if (WorkbookTemplate.IncludeProductName)
            {
                sizeOfSecondContentBlock++;

                targetWorksheet.Cells[currentRow, 1, currentRow, 2].Merge = true;
                targetWorksheet.Cells[currentRow, 1].Value = "Product";


                for (int i = 0; i < itemsInWorksheet; i++)
                {
                    string productCode = Lot.ProductCode(SalesOrder.Lots[(worksheetNumber - 1) * 6 + i]);

                    if (FinishedGoodsData.ProductNameExists(productCode, out string productName))
                    {
                        targetWorksheet.Cells[currentRow, 3 + i].Value = productName;
                    }
                    else
                    {
                        targetWorksheet.Cells[currentRow, 3 + i].Value = "Finished Goods Search Error";
                        targetWorksheet.Cells[currentRow, 3 + i].Style.Font.Color.SetColor(Color.Red);
                    }
                }
                currentRow++;
            }

            if (WorkbookTemplate.IncludeRecipeAndItem)
            {
                sizeOfSecondContentBlock++;

                targetWorksheet.Cells[currentRow, 1, currentRow, 2].Merge = true;
                targetWorksheet.Cells[currentRow, 1].Value = "Recipe/Item";

                for (int i = 0; i < itemsInWorksheet; i++)
                {
                    string productCode = Lot.ProductCode(SalesOrder.Lots[(worksheetNumber - 1) * 6 + i]);

                    if (FinishedGoodsData.RecipeCodeExists(productCode, out string recipeCode))
                    {
                        targetWorksheet.Cells[currentRow, 3 + i].Value = recipeCode + "/" + productCode;
                    }
                    else
                    {
                        targetWorksheet.Cells[currentRow, 3 + i].Value = "Finished Goods Search Error";
                        targetWorksheet.Cells[currentRow, 3 + i].Style.Font.Color.SetColor(Color.Red);
                    }

                }
                currentRow++;
            }

            if (WorkbookTemplate.IncludeLotCode)
            {
                sizeOfSecondContentBlock++;

                targetWorksheet.Cells[currentRow, 1, currentRow, 2].Merge = true;
                targetWorksheet.Cells[currentRow, 3, currentRow, 8].Style.Numberformat.Format = "0"; // Prevents Excel from converting to scientific notation
                targetWorksheet.Cells[currentRow, 1].Value = "Lot Code";

                string lotCode;
                long lotCodeAsLong;
                for (int i = 0; i < itemsInWorksheet; i++)
                {
                    if (WorkbookTemplate.SelectedAlgorithm == Templates.Template.Algorithm.Standard)
                    {
                        lotCode = SalesOrder.Lots[i + (worksheetNumber - 1) * 6];

                        if (Int64.TryParse(lotCode, out lotCodeAsLong) && lotCode.Length == 13)
                        {
                            targetWorksheet.Cells[currentRow, 3 + i].Value = lotCodeAsLong;
                        }
                        else if (lotCode.Length > 13)
                        {
                            targetWorksheet.Cells[currentRow, 3 + i].Value = "Lot too long";
                            targetWorksheet.Cells[currentRow, 3 + i].Style.Font.Color.SetColor(Color.Red);
                        }
                        else if (lotCode.Length < 13)
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
                }
                currentRow++;
            }

            if (WorkbookTemplate.IncludeBatchFromMicro)
            {
                sizeOfSecondContentBlock++;

                targetWorksheet.Cells[currentRow, 1, currentRow, 2].Merge = true;
                targetWorksheet.Cells[currentRow, 1].Value = "Batch";

                for (int i = 0; i < itemsInWorksheet; i++)
                {
                    if (microIndices.Count > 0)
                    {
                        if (microIndices[i].Count > 0)
                        {
                            string retrievedBatchValue = NWAData.BatchValuesFromMicroIndices(microIndices[i]); // Some facilities put production lines in batch column

                            bool containsQualifyingLetter = false; // Some facilities put production lines in batch column

                            foreach (char batchChar in retrievedBatchValue) // Batches contain only numbers and the letter v (vat)
                            {
                                if (char.IsLetter(batchChar) && batchChar.ToString().ToLower() != "v")
                                {
                                    containsQualifyingLetter = true;
                                }
                            }

                            if (containsQualifyingLetter == false && retrievedBatchValue != "")   // Output raw string if okay
                            {
                                targetWorksheet.Cells[currentRow, 3 + i].Value = retrievedBatchValue;
                            }
                            else if (retrievedBatchValue == "")
                            {
                                targetWorksheet.Cells[currentRow, 3 + i].Value = "N/A";
                            }
                            else                                    // Output modified, colored string if maybe not okay
                            {
                                targetWorksheet.Cells[currentRow, 3 + i].Value = "Batch (" + retrievedBatchValue +
                                    ") potentially invalid";
                                targetWorksheet.Cells[currentRow, 3 + i].Style.Font.Color.SetColor(Color.OrangeRed);
                            }

                        }
                        else
                        {
                            targetWorksheet.Cells[currentRow, 3 + i].Value = "Micro search error";
                            targetWorksheet.Cells[currentRow, 3 + i].Style.Font.Color.SetColor(Color.Red);
                        }
                    }
                    else
                    {
                        targetWorksheet.Cells[currentRow, 3 + i].Value = "No usable data";
                        targetWorksheet.Cells[currentRow, 3 + i].Style.Font.Color.SetColor(Color.Red);
                    }
                }
                currentRow++;
            }

            if (WorkbookTemplate.IncludeBatchFromDressing)
            {
                // sizeOfSecondContentBlock modified at end of block

                for (int i = 0; i < itemsInWorksheet; i++)
                {
                    if (titrationIndices.Count > 0)
                    {
                        if (titrationIndices[i].Count > 0)
                        {
                            string batchValue = NWAData.BatchValuesFromTitrationIndices(microIndices[i]);

                            if (batchValue != "")
                            {
                                targetWorksheet.Cells[currentRow, 3 + i].Value = batchValue;
                            }
                            else
                            {
                                targetWorksheet.Cells[currentRow, 3 + i].Value = "N/A";
                            }
                        }
                        else
                        {
                            targetWorksheet.Cells[currentRow, 3 + i].Value = "Micro search error";
                            targetWorksheet.Cells[currentRow, 3 + i].Style.Font.Color.SetColor(Color.Red);
                        }
                    }
                    else
                    {
                        targetWorksheet.Cells[currentRow, 3 + i].Value = "No usable data";
                        targetWorksheet.Cells[currentRow, 3 + i].Style.Font.Color.SetColor(Color.Red);
                    }
                }

                if (WorkbookTemplate.IncludeBatchFromDressing == false)
                {
                    sizeOfSecondContentBlock++; // If both are set to be included, batch from dressing will overwrite
                    currentRow++;
                }
            }

            if (WorkbookTemplate.IncludeBestByDate)
            {
                sizeOfSecondContentBlock++;

                targetWorksheet.Cells[currentRow, 1, currentRow, 2].Merge = true;
                targetWorksheet.Cells[currentRow, 1].Value = "Best By Date";

                for (int i = 0; i < itemsInWorksheet; i++)
                {
                    if (WorkbookTemplate.SelectedAlgorithm == Templates.Template.Algorithm.Standard)
                    {
                        if (Lot.TryParseBestBy(SalesOrder.Lots[(worksheetNumber - 1) * 6 + i], out DateTime bestByDate) == true)
                        {
                            targetWorksheet.Cells[currentRow, 3 + i].Value = bestByDate.ToShortDateString();
                        }
                        else
                        {
                            targetWorksheet.Cells[currentRow, 3 + i].Value = "Date conversion error";
                            targetWorksheet.Cells[currentRow, 3 + i].Style.Font.Color.SetColor(Color.Red);
                        }
                    }
                }

                currentRow++;
            }

            if (WorkbookTemplate.IncludeManufacturingDate)
            {
                sizeOfSecondContentBlock++;

                targetWorksheet.Cells[currentRow, 1, currentRow, 2].Merge = true;
                targetWorksheet.Cells[currentRow, 1].Value = "Manufacturing Date";

                if (WorkbookTemplate.SelectedAlgorithm == Templates.Template.Algorithm.Standard)
                {
                    for (int i = 0; i < itemsInWorksheet; i++)
                    {
                        targetWorksheet.Cells[currentRow, 3 + i].Value = FinishedGoodsData.GetMadeDate(SalesOrder.Lots[(worksheetNumber - 1) * 6 + i]).ToShortDateString();
                    }
                }
                currentRow++;
            }

            // Insert comments containing each product's made date to assist user investigation if the information should otherwise not be present
            // Comments are supposed to be placed along the first row of the second content block
            if (WorkbookTemplate.IncludeManufacturingDate == false)
            {
                if (WorkbookTemplate.SelectedAlgorithm == Templates.Template.Algorithm.Standard)
                {
                    for (int i = 0; i < itemsInWorksheet; i++)
                    {
                        string commentValue = "Made: ";
                        commentValue += FinishedGoodsData.GetMadeDate(SalesOrder.Lots[(worksheetNumber - 1) * 6 + i]).ToShortDateString();

                        targetWorksheet.Cells[11 + sizeOfFirstContentBlock + 2, 3 + i].AddComment(commentValue, "CoA Tool");
                        targetWorksheet.Cells[11 + sizeOfFirstContentBlock + 2, 3 + i].Comment.BackgroundColor = Color.LightBlue;
                        targetWorksheet.Cells[11 + sizeOfFirstContentBlock + 2, 3 + i].Comment.LineColor = Color.LightSkyBlue;
                        targetWorksheet.Cells[11 + sizeOfFirstContentBlock + 2, 3 + i].Comment.AutoFit = true;
                    }
                }
            }

            if (WorkbookTemplate.IncludeManufacturingSite)
            {
                sizeOfSecondContentBlock++;

                targetWorksheet.Cells[currentRow, 1, currentRow, 2].Merge = true;
                targetWorksheet.Cells[currentRow, 1].Value = "Manufacturing Site";

                for (int i = 0; i < itemsInWorksheet; i++)
                {
                    targetWorksheet.Cells[currentRow, 3 + i].Value = Lot.StateAndCityOfManufacture(SalesOrder.Lots[(worksheetNumber - 1) * 6 + i]);
                }

                currentRow++;
            }

            // For both first and second content blocks: left-hand cell descriptions, second content block header, border drawing

            int sumOfBlockRows = sizeOfFirstContentBlock + sizeOfSecondContentBlock;

            if (sumOfBlockRows > 0) // TODO: Review block
            {
                if (sizeOfSecondContentBlock > 0)
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
                    targetWorksheet.Cells[secondBlockHeaderRow, 3, secondBlockHeaderRow + sizeOfSecondContentBlock, 8].Style.HorizontalAlignment = ExcelHorizontalAlignment.Center;

                    CreateTableOfBorders(6, sizeOfSecondContentBlock, secondBlockHeaderRow + 1, 3, targetWorksheet);
                }

                targetWorksheet.Cells[11, 2, 10 + sumOfBlockRows, 2].Style.Font.Italic = true;
                targetWorksheet.Cells[11, 1, 10 + sumOfBlockRows, 2].Style.HorizontalAlignment = ExcelHorizontalAlignment.Right;
            }

            // For third content block

            if (sumOfBlockRows > 0)
                currentRow += 2; // An empty space between blocks and an empty space for the block's header if the total block size > 0
            else
                currentRow = 12; // Allows the third block's header to use row 11

            int sizeOfThirdContentBlock = 0;

            if (WorkbookTemplate.IncludeAcidity)
            {
                sizeOfThirdContentBlock++;

                targetWorksheet.Cells[currentRow, 1].Value = "% Acid (TA)";
                targetWorksheet.Cells[currentRow, 2].Value = AcidMethod;

                for (int i = 0; i < itemsInWorksheet; i++)
                {
                    if (titrationIndices.Count > 0)
                    {
                        if (titrationIndices[i].Count > 0)
                        {
                            float acidity;

                            if (NWAData.TitrationValueExists(titrationIndices[i], CSV.NWAData.TitrationOffset.Acidity, out acidity))
                            {
                                targetWorksheet.Cells[currentRow, 3 + i].Value = Math.Round(acidity, 2);
                            }
                            else
                            {
                                targetWorksheet.Cells[currentRow, 3 + i].Value = "Missing value";
                                targetWorksheet.Cells[currentRow, 3 + i].Style.Font.Color.SetColor(Color.Red);
                            }

                        }
                        else
                        {
                            targetWorksheet.Cells[currentRow, 3 + i].Value = "Search error";
                            targetWorksheet.Cells[currentRow, 3 + i].Style.Font.Color.SetColor(Color.Red);
                        }
                    }
                    else
                    {
                        targetWorksheet.Cells[currentRow, 3 + i].Value = "No usable data";
                        targetWorksheet.Cells[currentRow, 3 + i].Style.Font.Color.SetColor(Color.Red);
                    }
                }

                currentRow++;
            }

            if (WorkbookTemplate.IncludepH)
            {
                sizeOfThirdContentBlock++;

                targetWorksheet.Cells[currentRow, 1].Value = "pH";
                targetWorksheet.Cells[currentRow, 2].Value = pHMethod;

                for (int i = 0; i < itemsInWorksheet; i++)
                {
                    if (titrationIndices.Count > 0)
                    {

                        if (titrationIndices[i].Count > 0)
                        {
                            float pHValue;

                            if (NWAData.TitrationValueExists(titrationIndices[i], CSV.NWAData.TitrationOffset.pH, out pHValue))
                            {
                                targetWorksheet.Cells[currentRow, 3 + i].Value = Math.Round(pHValue, 2);
                            }
                            else
                            {
                                targetWorksheet.Cells[currentRow, 3 + i].Value = "Missing value";
                                targetWorksheet.Cells[currentRow, 3 + i].Style.Font.Color.SetColor(Color.Red);
                            }
                        }
                        else
                        {
                            targetWorksheet.Cells[currentRow, 3 + i].Value = "Search error";
                            targetWorksheet.Cells[currentRow, 3 + i].Style.Font.Color.SetColor(Color.Red);
                        }
                    }
                    else
                    {
                        targetWorksheet.Cells[currentRow, 3 + i].Value = "No usable data";
                        targetWorksheet.Cells[currentRow, 3 + i].Style.Font.Color.SetColor(Color.Red);
                    }
                }

                currentRow++;
            }

            if (WorkbookTemplate.IncludeViscosityCM)
            {
                sizeOfThirdContentBlock++;

                targetWorksheet.Cells[currentRow, 1].Value = "Viscosity cm";
                targetWorksheet.Cells[currentRow, 2].Value = ViscosityCMMethod;

                for (int i = 0; i < itemsInWorksheet; i++)
                {
                    if (titrationIndices.Count > 0)
                    {
                        if (titrationIndices[i].Count > 0)
                        {
                            float viscosityCM;

                            if (NWAData.TitrationValueExists(titrationIndices[i], CSV.NWAData.TitrationOffset.ViscosityCM, out viscosityCM))
                            {
                                targetWorksheet.Cells[currentRow, 3 + i].Value = Math.Round(viscosityCM, 2);
                            }
                            else
                            {
                                targetWorksheet.Cells[currentRow, 3 + i].Value = "Missing value";
                                targetWorksheet.Cells[currentRow, 3 + i].Style.Font.Color.SetColor(Color.Red);
                            }
                        }
                        else
                        {
                            targetWorksheet.Cells[currentRow, 3 + i].Value = "Search error";
                            targetWorksheet.Cells[currentRow, 3 + i].Style.Font.Color.SetColor(Color.Red);
                        }
                    }
                    else
                    {
                        targetWorksheet.Cells[currentRow, 3 + i].Value = "No usable data";
                        targetWorksheet.Cells[currentRow, 3 + i].Style.Font.Color.SetColor(Color.Red);
                    }
                }

                currentRow++;
            }

            if (WorkbookTemplate.IncludeViscosityCPS)
            {
                sizeOfThirdContentBlock++;

                targetWorksheet.Cells[currentRow, 1].Value = "Viscosity cps";
                targetWorksheet.Cells[currentRow, 2].Value = ViscosityCPSMethod;

                for (int i = 0; i < itemsInWorksheet; i++)
                {
                    if (titrationIndices.Count > 0)
                    {
                        if (titrationIndices[i].Count > 0)
                        {
                            float viscosityCPS;

                            if (NWAData.TitrationValueExists(titrationIndices[i], CSV.NWAData.TitrationOffset.ViscosityCPS, out viscosityCPS))
                            {
                                targetWorksheet.Cells[currentRow, 3 + i].Value = viscosityCPS;
                                targetWorksheet.Cells[currentRow, 3 + i].Style.Numberformat.Format = "#,0";
                            }
                            else
                            {
                                targetWorksheet.Cells[currentRow, 3 + i].Value = "Missing value";
                                targetWorksheet.Cells[currentRow, 3 + i].Style.Font.Color.SetColor(Color.Red);
                            }
                        }
                        else
                        {
                            targetWorksheet.Cells[currentRow, 3 + i].Value = "Search error";
                            targetWorksheet.Cells[currentRow, 3 + i].Style.Font.Color.SetColor(Color.Red);
                        }
                    }
                    else
                    {
                        targetWorksheet.Cells[currentRow, 3 + i].Value = "No usable data";
                        targetWorksheet.Cells[currentRow, 3 + i].Style.Font.Color.SetColor(Color.Red);
                    }
                }

                currentRow++;
            }

            if (WorkbookTemplate.IncludeWaterActivity)
            {
                sizeOfThirdContentBlock++;

                targetWorksheet.Cells[currentRow, 1].Value = "Water activity (aW)";

                for (int i = 0; i < itemsInWorksheet; i++)
                {
                    if (titrationIndices.Count > 0)
                    {
                        if (titrationIndices[i].Count > 0)
                        {
                            bool waterActivityValuePresent = NWAData.WaterActivityExists(titrationIndices[i], out float waterActivity, out bool valueIsAlsoValid);

                            if (waterActivityValuePresent && valueIsAlsoValid)
                            {
                                targetWorksheet.Cells[currentRow, 3 + i].Value = Math.Round(waterActivity, 3);
                            }
                            else if (waterActivityValuePresent && valueIsAlsoValid == false)
                            {
                                targetWorksheet.Cells[currentRow, 3 + i].Value = "Invalid value";
                                targetWorksheet.Cells[currentRow, 3 + i].Style.Font.Color.SetColor(Color.Red);
                            }
                            else
                            {
                                targetWorksheet.Cells[currentRow, 3 + i].Value = "Missing value";
                                targetWorksheet.Cells[currentRow, 3 + i].Style.Font.Color.SetColor(Color.Red);
                            }
                        }
                        else
                        {
                            targetWorksheet.Cells[currentRow, 3 + i].Value = "Search error";
                            targetWorksheet.Cells[currentRow, 3 + i].Style.Font.Color.SetColor(Color.Red);
                        }
                    }
                    else
                    {
                        targetWorksheet.Cells[currentRow, 3 + i].Value = "No usable data";
                        targetWorksheet.Cells[currentRow, 3 + i].Style.Font.Color.SetColor(Color.Red);
                    }
                }

                currentRow++;
            }

            if (WorkbookTemplate.IncludeBrixSlurry)
            {
                sizeOfThirdContentBlock++;

                targetWorksheet.Cells[currentRow, 1].Value = "Brix slurry";

                currentRow++;
            }

            if (sizeOfThirdContentBlock > 0) // TODO: Review block
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

                targetWorksheet.Cells[thirdBlockHeaderRow, 3, thirdBlockHeaderRow + sizeOfThirdContentBlock, 8].Style.HorizontalAlignment = ExcelHorizontalAlignment.Center;

                targetWorksheet.Cells[thirdBlockHeaderRow, 3].Value = "Analytical Results";

                targetWorksheet.Cells[thirdBlockHeaderRow, 3, thirdBlockHeaderRow, 8].Merge = true;

                targetWorksheet.Cells[thirdBlockHeaderRow, 1, thirdBlockHeaderRow, 3].Style.Font.Bold = true;

                targetWorksheet.Cells[thirdBlockHeaderRow, 1, thirdBlockHeaderRow, 8].Style.Border.Bottom.Style = ExcelBorderStyle.Medium;

                targetWorksheet.Cells[thirdBlockHeaderRow + 1, 2, thirdBlockHeaderRow + sizeOfThirdContentBlock, 2].Style.Font.Size = 6;

                targetWorksheet.Cells[thirdBlockHeaderRow + 1, 2, thirdBlockHeaderRow + sizeOfThirdContentBlock, 2].Style.HorizontalAlignment = ExcelHorizontalAlignment.Right;

                targetWorksheet.Cells[thirdBlockHeaderRow + 1, 1, thirdBlockHeaderRow + sizeOfThirdContentBlock, 1].Style.HorizontalAlignment = ExcelHorizontalAlignment.Left;
            }

            // For fourth content block

            if (sumOfBlockRows > 0)
            {
                currentRow += 2; // An empty space between blocks and an empty space for the block's header if the total block size > 0
            }
            else
            {
                currentRow = 12; // Allows the third block's header to use row 11
            }

            int sizeOfFourthContentBlock = 0;

            if (WorkbookTemplate.IncludeYeast)
            {
                sizeOfFourthContentBlock++;

                targetWorksheet.Cells[currentRow, 1].Value = "Yeast";
                targetWorksheet.Cells[currentRow, 2].Value = YeastMethod;

                for(int columnIterator = 0; columnIterator < itemsInWorksheet; columnIterator++)
                {
                    PopulateMicroCell(CSV.NWAData.MicroOffset.Yeast, targetWorksheet, currentRow, columnIterator, microIndices[columnIterator]);
                }

                currentRow++;
            }

            if (WorkbookTemplate.IncludeMold)
            {
                sizeOfFourthContentBlock++;

                targetWorksheet.Cells[currentRow, 1].Value = "Mold";
                targetWorksheet.Cells[currentRow, 2].Value = MoldMethod;

                for (int columnIterator = 0; columnIterator < itemsInWorksheet; columnIterator++)
                {
                    PopulateMicroCell(CSV.NWAData.MicroOffset.Mold, targetWorksheet, currentRow, columnIterator, microIndices[columnIterator]);
                }

                currentRow++;
            }

            if (WorkbookTemplate.IncludeAerobic)
            {
                sizeOfFourthContentBlock++;

                targetWorksheet.Cells[currentRow, 1].Value = "Aerobic";
                targetWorksheet.Cells[currentRow, 2].Value = AerobicMethod;

                for (int columnIterator = 0; columnIterator < itemsInWorksheet; columnIterator++)
                {
                    PopulateMicroCell(CSV.NWAData.MicroOffset.Aerobic, targetWorksheet, currentRow, columnIterator, microIndices[columnIterator]);
                }

                currentRow++;
            }

            if (WorkbookTemplate.IncludeColiform)
            {
                sizeOfFourthContentBlock++;

                targetWorksheet.Cells[currentRow, 1].Value = "Total coliform";
                targetWorksheet.Cells[currentRow, 2].Value = ColiformMethod;

                for (int columnIterator = 0; columnIterator < itemsInWorksheet; columnIterator++)
                {
                    PopulateMicroCell(CSV.NWAData.MicroOffset.Coliform, targetWorksheet, currentRow, columnIterator, microIndices[columnIterator]);
                }

                currentRow++;
            }

            if (WorkbookTemplate.IncludeEColi)
            {
                sizeOfFourthContentBlock++;

                targetWorksheet.Cells[currentRow, 1].Value = "E. coliform";
                targetWorksheet.Cells[currentRow, 2].Value = EColiMethod;

                for (int columnIterator = 0; columnIterator < itemsInWorksheet; columnIterator++)
                {
                    PopulateMicroCell(CSV.NWAData.MicroOffset.EColi, targetWorksheet, currentRow, columnIterator, microIndices[columnIterator]);
                }

                currentRow++;
            }

            if (WorkbookTemplate.IncludeLactics)
            {
                sizeOfFourthContentBlock++;

                targetWorksheet.Cells[currentRow, 1].Value = "Lactics";
                targetWorksheet.Cells[currentRow, 2].Value = LacticMethod;

                for (int columnIterator = 0; columnIterator < itemsInWorksheet; columnIterator++)
                {
                    PopulateMicroCell(CSV.NWAData.MicroOffset.Lactic, targetWorksheet, currentRow, columnIterator, microIndices[columnIterator]);
                }

                currentRow++;
            }

            if (WorkbookTemplate.IncludeSalmonella)
            {
                sizeOfFourthContentBlock++;

                targetWorksheet.Cells[currentRow, 1].Value = "Salmonella";

                currentRow++;
            }

            if (WorkbookTemplate.IncludeListeria)
            {
                sizeOfFourthContentBlock++;

                targetWorksheet.Cells[currentRow, 1].Value = "Listeria";

                currentRow++;
            }

            if (sizeOfFourthContentBlock > 0) // TODO: Review block
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

                targetWorksheet.Cells[fourthBlockHeaderRow, 3, fourthBlockHeaderRow + sizeOfFourthContentBlock, 8].Style.HorizontalAlignment = ExcelHorizontalAlignment.Center;

                targetWorksheet.Cells[fourthBlockHeaderRow, 3, fourthBlockHeaderRow + sizeOfFourthContentBlock, 8].Style.Numberformat.Format = "#,##0";

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

            if (WorkbookTemplate.IncludeColorAndAppearance)
            {
                sizeOfFifthContentBlock++;

                targetWorksheet.Cells[currentRow, 1].Value = "Color/Appearance";

                currentRow++;
            }

            if (WorkbookTemplate.IncludeForm)
            {
                sizeOfFifthContentBlock++;

                targetWorksheet.Cells[currentRow, 1].Value = "Form";

                currentRow++;
            }

            if (WorkbookTemplate.IncludeFlavorAndOdor)
            {
                sizeOfFifthContentBlock++;

                targetWorksheet.Cells[currentRow, 1].Value = "Flavor/Odor";

                currentRow++;
            }

            if (sizeOfFifthContentBlock > 0) // TODO: Review block
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
        /// <summary>
        /// Determines the string value of a cell for micro results in the worksheet.
        /// </summary>
        /// <param name="rawMicroValues"></param>
        /// <param name="microType"></param>
        /// <returns></returns>
        private string ValueForMicroCell(List<string> rawMicroValues, CSV.NWAData.MicroOffset microType, string cityOfManufacture)
        {
            if(rawMicroValues.Count == 0)
            {
                return "Search error";
            }

            List<int> numericalMicroValues = OnlyIntegersFrom(rawMicroValues);

            if(numericalMicroValues.Count > 0)
            {
                int greatestValue = 0;

                foreach (int value in numericalMicroValues)
                {
                    if (value > greatestValue)
                    {
                        greatestValue = value;
                    }
                }

                if(greatestValue == 0)
                {
                    return XMLOperations.Definitions.NodeValueViaXPath(XMLOperations.Definitions.DefinitionFiles.Micro, 
                        "/Dilutions_By_Factory/factory[@name = '" + cityOfManufacture + "']/" + Convert.ToString(microType));
                }
                else

                {
                    return Convert.ToString(greatestValue);
                }
            }
            else
            {
                return "N/A";
            }
        }
        /// <summary>
        /// Returns the converted values of integer-convertable strings.
        /// </summary>
        /// <param name="targetValues">The strings to attempt integer conversion with.</param>
        /// <returns></returns>
        private List<int> OnlyIntegersFrom(List<string> targetValues)
        {
            List<int> integerParseableValues = new List<int>();

            foreach(string item in targetValues)
            {
                if(Int32.TryParse(item, out int result)) // Result is left unused
                {
                    integerParseableValues.Add(result);
                }
            }

            return integerParseableValues;
        }
        /// <summary>
        /// Returns a Color dependent on the target cell's value and comment.
        /// </summary>
        /// <param name="targetWorksheet">The worksheet object containing the target cell.</param>
        /// <param name="targetCellRow">Represents the row in which the target cell is positioned.</param>
        /// <param name="targetCellColumn">Represents the column in which the target cell is positioned.</param>
        /// <returns></returns>
        private Color FontColorForCell(ExcelWorksheet targetWorksheet, int targetCellRow, int targetCellColumn, string commentText)
        {
           string cellValue = targetWorksheet.Cells[targetCellRow, targetCellColumn].Value.ToString();

            if (cellValue == "Search error")
            {
                return Color.Red;
            }
            else if (commentText.Contains("invalid"))
            {
                return Color.OrangeRed;
            }
            else
            {
                return Color.Black;
            }
        }
        /// <summary>
        /// Returns a Color dependent on the target cell's value.
        /// </summary>
        /// <param name="targetWorksheet">The worksheet object containing the target cell.</param>
        /// <param name="targetCellRow">Represents the row in which the target cell is positioned.</param>
        /// <param name="targetCellColumn">Represents the column in which the target cell is positioned.</param>
        /// <returns></returns>
        private Color FontColorForCell(ExcelWorksheet targetWorksheet, int targetCellRow, int targetCellColumn)
        {
            string cellValue = targetWorksheet.Cells[targetCellRow, targetCellColumn].Value.ToString();

            if (cellValue == "Search error")
            {
                return Color.Red;
            }
            else
            {
                return Color.Black;
            }
        }
        /// <summary>
        /// Determines if a comment is required for a particular micro cell.
        /// </summary>
        /// <param name="rawMicroValues">The micro values relevant to the cell.</param>
        /// <param name="commentText">The value of the comment to be written.</param>
        /// <returns></returns>
        private bool CommentNeededForMicroCell(List<string> rawMicroValues, out string commentText)
        {
            foreach(string rawValue in rawMicroValues)
            {
                if(rawValue != "*" && Int32.TryParse(rawValue, out _ ) == false)
                {
                    commentText = "One or more values were invalid";
                    return true;
                }
            }
            commentText = "";
            return false;
        }
        /// <summary>
        /// Assigns the target cell with relevant information and styling pertaining to micro results.
        /// </summary>
        /// <param name="microTypeOfCell">Indentifies the type of micro result the cell is associated with.</param>
        /// <param name="targetWorksheet">The worksheet containing the target cell.</param>
        /// <param name="cellRow">The row in which the target cell is positioned.</param>
        /// <param name="cellColumn">The column in which the target cell is positioned.</param>
        /// <param name="microIndicesForNWA">The micro result indices for values relevant to the cell.</param>
        private void PopulateMicroCell(CSV.NWAData.MicroOffset microTypeOfCell, ExcelWorksheet targetWorksheet, int cellRow, int cellColumn, List<int> microIndicesForNWA)
        {
            // Get all relevant values
            List<string> rawMicroValues = NWAData.MicroValues(microIndicesForNWA, microTypeOfCell);

            // Determine manufacturing city for parsing Micro.XML
            string cityOfManufacture = Lot.CityOfManufacture(SalesOrder.Lots[(Convert.ToInt32(targetWorksheet.Name.Substring(4)) - 1) * 6 + cellColumn]);
            // Assign final value to cell
            string valueForCell = ValueForMicroCell(rawMicroValues, microTypeOfCell, cityOfManufacture);
            targetWorksheet.Cells[cellRow, 3 + cellColumn].Value = valueForCell;

            // Reassign the cell value from a string to an integer if integer conversion succeeds
            if (Int32.TryParse(valueForCell, out int parsedValue))
            {
                // Excel throws warning for numbers input as strings regardless of cell formatting.
                targetWorksheet.Cells[cellRow, 3 + cellColumn].Value = parsedValue;
            }

            // Set cell font color and add a comment if necessary
            Color fontColor;
            if (CommentNeededForMicroCell(rawMicroValues, out string commentText))
            {
                targetWorksheet.Cells[cellRow, 3 + cellColumn].AddComment(commentText, "CoA Tool");
                fontColor = FontColorForCell(targetWorksheet, cellRow, 3 + cellColumn, commentText);
            }
            else
            {
                fontColor = FontColorForCell(targetWorksheet, cellRow, 3 + cellColumn);
            }
            targetWorksheet.Cells[cellRow, 3 + cellColumn].Style.Font.Color.SetColor(fontColor);
        }
        
    }
}

