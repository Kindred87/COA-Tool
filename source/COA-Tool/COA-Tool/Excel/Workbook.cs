using System;
using System.Collections.Generic;
using System.Text;
using OfficeOpenXml;
using System.Xml;
using System.Drawing;
using OfficeOpenXml.Style;
using System.IO;
using OfficeOpenXml.Drawing;


namespace COA_Tool.Excel
{
    class Workbook
    {
        public enum CustomerName { TaylorFarmsTennessee, Latitute36 };
        private enum TitrationOffset { Acidity = 5, Viscosity = 2, Salt = 4, pH = 6 }
        private enum MicroOffset { Yeast = 9, Mold = 11, Aerobic = 15, Coliform = 7, Lactic = 13 }
        private CustomerName CustomerNameForWorkbook;

        private string AcidMethod = "(AOAC 30.048 14th Ed.)  ";
        private string pHMethod = "(AOAC 30.012 14th Ed.)  ";
        private string ViscosityCPSMethod = "(Brookfield)  ";
        private string SaltMethod = "(AOAC 937.09 18th Ed.)  ";
        private string YeastMethod = "(AOAC 997.02)  ";
        private string MoldMethod = "(AOAC 997.02)  ";
        private string AerobicMethod = "(AOAC 990.12)  ";
        private string ColiformMethod = "(AOAC 991.14)  ";
        private string LacticMethod = "(AOAC 990.12)  ";

        private List<List<string>> TableauData;
        
        private List<List<string>> TitrationResults;
        private List<List<string>> MicroResults;
        private List<List<string>> FinishedGoods;
        private List<List<string>> Recipes;
        public Workbook(List<List<string>> tableauData, List<List<string>> titrationResults, List<List<string>> microResults, CustomerName name,
            List<List<string>> finishedGoods, List<List<string>> recipes)
        {
            CustomerNameForWorkbook = name;
            TableauData = tableauData;
            TitrationResults = titrationResults;
            MicroResults = microResults;
            FinishedGoods = finishedGoods;
            Recipes = recipes;
        }

        public void Generate()
        {
            //Thread.CurrentThread.IsBackground = false;

            using (ExcelPackage package = new ExcelPackage())
            {
                int pageCount = 0;

                if (CustomerNameForWorkbook == CustomerName.TaylorFarmsTennessee)
                {
                    pageCount = (TableauData.Count - 1) / 4;
                    if ((TableauData.Count - 1) % 4 > 0)
                        pageCount++;
                }
                else if (CustomerNameForWorkbook == CustomerName.Latitute36)
                {
                    pageCount = (TableauData.Count - 1) / 6;
                    if ((TableauData.Count - 1) % 6 > 0)
                        pageCount++;
                }

                for (int i = 1; i <= pageCount; i++)
                {
                    ExcelWorksheet worksheet = package.Workbook.Worksheets.Add("Page" + i);

                    PopulateUniversalContent(worksheet);
                    PopulateContentsByCustomer(worksheet, i);
                }


                if (Directory.Exists(Environment.GetFolderPath(Environment.SpecialFolder.Desktop) + "/COAs") == false)
                {
                    Directory.CreateDirectory(Environment.GetFolderPath(Environment.SpecialFolder.Desktop) + "/COAs");
                }
                package.SaveAs(new FileInfo(Environment.GetFolderPath(Environment.SpecialFolder.Desktop) + "/COAs/" + TableauData[1][3] + ".xlsx"));
            }
        }
        /// <summary>
        /// Creates formatting and content shared by all document designs
        /// </summary>
        /// <param name="worksheet"></param>
        private void PopulateUniversalContent(ExcelWorksheet worksheet)
        {
            worksheet.Cells.Style.Font.SetFromFont(new Font("Calibri", 11));



            worksheet.View.ShowGridLines = false;
        }
        /// <summary>
        /// Determines which customer-specific worksheet design to use and calls the appropriate method
        /// </summary>
        /// <param name="worksheet"></param>
        private void PopulateContentsByCustomer(ExcelWorksheet worksheet, int page)
        {
            if (CustomerNameForWorkbook == CustomerName.TaylorFarmsTennessee)
            {
                PopulateContentsTaylorFarmTennessee(worksheet, page);
            }
            else if (CustomerNameForWorkbook == CustomerName.Latitute36)
            {
                PopulateContentsLatitute36(worksheet, page);
            }
        }
        private void PopulateContentsTaylorFarmTennessee(ExcelWorksheet worksheet, int page)
        {

            worksheet.Column(1).Width = 20;
            worksheet.Column(2).Width = 12;
            worksheet.Column(3).Width = 16;
            worksheet.Column(4).Width = 15.72;
            worksheet.Column(5).Width = 15.72;
            worksheet.Column(6).Width = 15.72;

            Image image = Image.FromFile("LH logo.png");
            ExcelPicture logo = worksheet.Drawings.AddPicture("Logo", image);
            logo.SetSize(57);
            logo.SetPosition(10, 330);

            CreateTableOfBorders(4, 4, 11, 3, worksheet);
            CreateTableOfBorders(4, 5, 16, 3, worksheet);
            CreateTableOfBorders(4, 4, 22, 3, worksheet);
            CreateTableOfBorders(4, 1, 30, 3, worksheet);

            worksheet.Cells["C8"].Value = "Certificate of Analysis";
            worksheet.Cells["C8"].Style.Font.SetFromFont(new Font("Calibri", 26, FontStyle.Bold));
            worksheet.Cells["C8"].Style.HorizontalAlignment = ExcelHorizontalAlignment.Center;
            worksheet.Cells["C8:F9"].Merge = true;

            worksheet.Cells["C11"].Value = "Taylor Farms / Tennessee";
            worksheet.Cells["C11"].Style.Font.Size = 12;
            worksheet.Cells["C11"].Style.Font.Bold = true;
            worksheet.Cells["C11"].Style.HorizontalAlignment = ExcelHorizontalAlignment.Center;
            worksheet.Cells["C11:F11"].Merge = true;

            worksheet.Cells["C12"].Value = TableauData[1][3];
            worksheet.Cells["C12"].Style.Font.Size = 12;
            worksheet.Cells["C12"].Style.Font.Bold = true;
            worksheet.Cells["C12"].Style.HorizontalAlignment = ExcelHorizontalAlignment.Center;
            worksheet.Cells["C12:F12"].Merge = true;

            worksheet.Cells["C13"].Style.Font.Size = 12;
            worksheet.Cells["C13"].Style.Font.Bold = true;
            worksheet.Cells["C13"].Style.HorizontalAlignment = ExcelHorizontalAlignment.Center;
            worksheet.Cells["C13:F13"].Merge = true;

            worksheet.Cells["C14"].Value = DateTime.Now.Date.ToShortDateString();
            worksheet.Cells["C14"].Style.Font.Size = 12;
            worksheet.Cells["C14"].Style.Font.Bold = true;
            worksheet.Cells["C14"].Style.HorizontalAlignment = ExcelHorizontalAlignment.Center;
            worksheet.Cells["C14:F14"].Merge = true;

            worksheet.Cells["A11"].Value = "Customer:";
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
            worksheet.Cells["A30"].Value = "Verified By";
            worksheet.Cells["C30:F30"].Merge = true;
            worksheet.Cells["C30"].Style.HorizontalAlignment = ExcelHorizontalAlignment.Center;

            worksheet.Cells["A32:F36"].Merge = true;
            worksheet.Cells["A32"].Style.WrapText = true;
            worksheet.Cells["A32"].Value = "Please be advised that the results herein are accurate and representative to the best of our knowledge, " +
                "based on current data and information as of the date of this document.  This COA is intended only for the person or company specified " +
                "hereon as the recipient.  This COA shall not be distributed to any person or company other than the intended recipient.  A person or " +
                "company other than the one named hereon shall not rely or make use of the statements and results herein. ";

            for (int i = 1; i <= 4; i++)
            {
                string lotCode = GetLotCode(i + (page - 1) * 4);

                if (string.IsNullOrEmpty(lotCode))
                {
                    //lotCode = GetLotCode(i + 1 + (page - 1) * 4);
                    //if (string.IsNullOrEmpty(lotCode))
                        continue;
                }

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
        private void PopulateContentsLatitute36(ExcelWorksheet worksheet, int page)
        {
            worksheet.Column(1).Width = 21;
            worksheet.Column(2).Width = 13;
            worksheet.Column(3).Width = 15.72;
            worksheet.Column(4).Width = 15.72;
            worksheet.Column(5).Width = 15.72;
            worksheet.Column(6).Width = 15.72;
            worksheet.Column(7).Width = 15.72;
            worksheet.Column(8).Width = 15.72;

            Image image = Image.FromFile("LH logo.png");
            ExcelPicture logo = worksheet.Drawings.AddPicture("Logo", image);
            logo.SetSize(57);
            logo.SetPosition(10, 445);

            CreateTableOfBorders(6, 4, 11, 3, worksheet);
            CreateTableOfBorders(6, 4, 16, 3, worksheet);
            CreateTableOfBorders(6, 8, 21, 3, worksheet);
            CreateTableOfBorders(6, 1, 30, 3, worksheet);

            worksheet.Cells["C8"].Value = "Certificate of Analysis";
            worksheet.Cells["C8"].Style.Font.SetFromFont(new Font("Calibri", 26, FontStyle.Bold));
            worksheet.Cells["C8"].Style.HorizontalAlignment = ExcelHorizontalAlignment.Center;
            worksheet.Cells["C8:H9"].Merge = true;

            worksheet.Cells["C11"].Value = "Latitute 36 (Ohio)";
            worksheet.Cells["C11"].Style.Font.Size = 12;
            worksheet.Cells["C11"].Style.Font.Bold = true;
            worksheet.Cells["C11"].Style.HorizontalAlignment = ExcelHorizontalAlignment.Center;
            worksheet.Cells["C11:H11"].Merge = true;

            worksheet.Cells["C12"].Value = Convert.ToInt32(TableauData[1][3]);
            worksheet.Cells["C12"].Style.Numberformat.Format = "0";
            worksheet.Cells["C12"].Style.Font.Size = 12;
            worksheet.Cells["C12"].Style.Font.Bold = true;
            worksheet.Cells["C12"].Style.HorizontalAlignment = ExcelHorizontalAlignment.Center;
            worksheet.Cells["C12:H12"].Merge = true;

            worksheet.Cells["C13"].Style.Font.Size = 12;
            worksheet.Cells["C13"].Style.Numberformat.Format = "0";
            worksheet.Cells["C13"].Style.Font.Bold = true;
            worksheet.Cells["C13"].Style.HorizontalAlignment = ExcelHorizontalAlignment.Center;
            worksheet.Cells["C13:H13"].Merge = true;

            worksheet.Cells["C14"].Value = DateTime.Now.Date.ToShortDateString();
            worksheet.Cells["C14"].Style.Numberformat.Format = "m/d/yyyy";
            worksheet.Cells["C14"].Style.Font.Size = 12;
            worksheet.Cells["C14"].Style.Font.Bold = true;
            worksheet.Cells["C14"].Style.HorizontalAlignment = ExcelHorizontalAlignment.Center;
            worksheet.Cells["C14:H14"].Merge = true;

            worksheet.Cells["A11"].Value = "Customer:";
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
            worksheet.Cells["B21"].Style.HorizontalAlignment = ExcelHorizontalAlignment.Right;
            worksheet.Cells["A22"].Value = "pH";
            worksheet.Cells["B22"].Value = pHMethod;
            worksheet.Cells["B22"].Style.Font.Size = 6;
            worksheet.Cells["B22"].Style.HorizontalAlignment = ExcelHorizontalAlignment.Right;
            worksheet.Cells["A23"].Value = "Viscosity cps";
            worksheet.Cells["B23"].Value = ViscosityCPSMethod;
            worksheet.Cells["B23"].Style.Font.Size = 6;
            worksheet.Cells["B23"].Style.HorizontalAlignment = ExcelHorizontalAlignment.Right;
            worksheet.Cells["A24"].Value = "Yeast cfu/gram";
            worksheet.Cells["B24"].Value = YeastMethod;
            worksheet.Cells["B24"].Style.Font.Size = 6;
            worksheet.Cells["B24"].Style.HorizontalAlignment = ExcelHorizontalAlignment.Right;
            worksheet.Cells["A25"].Value = "Mold cfu/gram";
            worksheet.Cells["B25"].Value = MoldMethod;
            worksheet.Cells["B25"].Style.Font.Size = 6;
            worksheet.Cells["B25"].Style.HorizontalAlignment = ExcelHorizontalAlignment.Right;
            worksheet.Cells["A26"].Value = "Aerobic cfu/gram";
            worksheet.Cells["B26"].Value = AerobicMethod;
            worksheet.Cells["B26"].Style.Font.Size = 6;
            worksheet.Cells["B26"].Style.HorizontalAlignment = ExcelHorizontalAlignment.Right;
            worksheet.Cells["A27"].Value = "Total coliform cfu/gram";
            worksheet.Cells["B27"].Value = ColiformMethod;
            worksheet.Cells["B27"].Style.Font.Size = 6;
            worksheet.Cells["B27"].Style.HorizontalAlignment = ExcelHorizontalAlignment.Right;
            worksheet.Cells["A28"].Value = "Lactics cfu/gram";
            worksheet.Cells["B28"].Value = LacticMethod;
            worksheet.Cells["B28"].Style.Font.Size = 6;
            worksheet.Cells["B28"].Style.HorizontalAlignment = ExcelHorizontalAlignment.Right;

            worksheet.Cells["A30"].Value = "Verified By";
            worksheet.Cells["C30:H30"].Merge = true;
            worksheet.Cells["C30"].Style.HorizontalAlignment = ExcelHorizontalAlignment.Center;

            worksheet.Cells["A32:F36"].Merge = true;
            worksheet.Cells["A32"].Style.WrapText = true;
            worksheet.Cells["A32"].Value = "Please be advised that the results herein are accurate and representative to the best of our knowledge, " +
                "based on current data and information as of the date of this document.  This COA is intended only for the person or company specified " +
                "hereon as the recipient.  This COA shall not be distributed to any person or company other than the intended recipient.  A person or " +
                "company other than the one named hereon shall not rely or make use of the statements and results herein. ";

            for (int i = 1; i <= 6; i++)
            {
                string lotCode = GetLotCode(i + (page - 1) * 6);

                if (string.IsNullOrEmpty(lotCode))
                {
                    lotCode = GetLotCode(i + 1 + (page - 1) * 6);
                    if (string.IsNullOrEmpty(lotCode))
                        continue;
                }

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

                worksheet.Cells[29, 1].Value = "Micro Index";
                worksheet.Cells[31, 1].Value = "Titration Index";
                foreach (int index in microIndices)
                {
                    worksheet.Cells[29, 2 + i].Value += Convert.ToString(index + 1) + ", ";
                }
                foreach (int index in titrationIndices)
                {
                    worksheet.Cells[31, 2 + i].Value += Convert.ToString(index + 1) + ", ";
                }

                if (microIndices.Count == 0)
                {
                    worksheet.Cells[24, 2 + i, 28, 2 + i].Value = "No result";
                    worksheet.Cells[24, 2 + i, 28, 2 + i].Style.Font.Color.SetColor(Color.Red);
                }
                else
                {
                    string factoryCode = GetFactoryCode(lotCode);

                    int yeastCount = GetMicroValue(microIndices, MicroOffset.Yeast);
                    //worksheet.Cells[24, 2 + i].Style.HorizontalAlignment = ExcelHorizontalAlignment.Center;
                    if (yeastCount == 0)
                        worksheet.Cells[24, 2 + i].Value = "<10";
                    else if (yeastCount > 0)
                        worksheet.Cells[24, 2 + i].Value = yeastCount;
                    else if (yeastCount == -1)
                        worksheet.Cells[24, 2 + i].Value = "N/A";

                    worksheet.Cells[24, 2 + i].Style.Numberformat.Format = "#,##0";


                    int moldCount = GetMicroValue(microIndices, MicroOffset.Mold);
                    //worksheet.Cells[25, 2 + i].Style.HorizontalAlignment = ExcelHorizontalAlignment.Center;
                    if (moldCount == 0)
                        worksheet.Cells[25, 2 + i].Value = "<10";
                    else if (moldCount > 0)
                        worksheet.Cells[25, 2 + i].Value = moldCount;
                    else if (moldCount == -1)
                        worksheet.Cells[25, 2 + i].Value = "N/A";

                    worksheet.Cells[25, 2 + i].Style.Numberformat.Format = "#,##0";

                    int aerobicCount = GetMicroValue(microIndices, MicroOffset.Aerobic);
                    //worksheet.Cells[26, 2 + i].Style.HorizontalAlignment = ExcelHorizontalAlignment.Center;
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
                    //worksheet.Cells[27, 2 + i].Style.HorizontalAlignment = ExcelHorizontalAlignment.Center;
                    if (coliformCount == 0)
                        worksheet.Cells[27, 2 + i].Value = "<10";
                    else if (coliformCount > 0)
                        worksheet.Cells[27, 2 + i].Value = coliformCount;
                    else if (coliformCount == -1)
                        worksheet.Cells[27, 2 + i].Value = "N/A";

                    worksheet.Cells[27, 2 + i].Style.Numberformat.Format = "#,##0";

                    int lacticCount = GetMicroValue(microIndices, MicroOffset.Lactic);
                    //worksheet.Cells[28, 2 + i].Style.HorizontalAlignment = ExcelHorizontalAlignment.Center;
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
            string madeDateAsTwoDigitYearString = madeDate.ToString("M/dd/yy");

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

            foreach (int value in microValues)
            {
                if (value > largestValue)
                {
                    largestValue = value;
                }
            }

            return largestValue;
        }
        /// <summary>
        /// Retrieves shelf life for a given recipe from the Recipes list
        /// </summary>
        /// <param name="recipeCode"></param>
        /// <returns></returns>
        private int DaysToExpiryForRecipe(string recipeCode)
        {
            foreach (List<string> row in Recipes)
            {
                if (row[0] == recipeCode)
                    return Convert.ToInt32(row[1]);
            }

            return -1;
        }
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
    }
}
