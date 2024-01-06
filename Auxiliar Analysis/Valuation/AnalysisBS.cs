using OfficeOpenXml.Style;
using OfficeOpenXml;
using System.Drawing;
using System.Globalization;
using Mayntech___Individual_Solution.Auxiliar.Auxiliar.FMP;
using Mayntech___Individual_Solution.Auxiliar.Auxiliar.Other;

namespace Mayntech___Individual_Solution.Auxiliar.Analysis
{
    public class AnalysisBS : ExcelNextCol
    {
        public async Task CreateBSAnalysis(ExcelPackage package, List<FinancialStatements> BalanceSheet, List<FinancialStatements> incomestatement, int col, int row, string companyName, int numberOfYears)
        {
            var workSheet = package.Workbook.Worksheets.Add("BS - Analysis");
            workSheet.View.ShowGridLines = false;
            workSheet.View.ZoomScale = 80;

            List<double> revenueValues = new List<double>();
            foreach (FinancialStatements item in BalanceSheet)
            {
                revenueValues.Add(item.Revenue);
            }

            // Cria a coluna azul em cima da tabela
            workSheet.Cells[2, 2].Value = "BS Analysis - " + companyName;
            workSheet.Cells[2, 2].Style.Font.Bold = true;
            workSheet.Cells[2, 2].Style.Font.Color.SetColor(Color.White);


            // atenção que este "H2" está hardcoded. Tem de ser refeito para variar com o número de anos
            workSheet.Cells["B2"].Style.Fill.PatternType = ExcelFillStyle.Solid;
            workSheet.Cells["B2"].Style.Fill.BackgroundColor.SetColor(Cores.CorPrincipal);

            workSheet.Cells[3, 2].Formula = "='P&L'!B3";
            workSheet.Cells[3, 2].Style.Font.Bold = true;
            workSheet.Cells[3, 2].Style.Font.Color.SetColor(Cores.CorTexto);
            workSheet.Cells[3, 2].Style.Border.Bottom.Style = ExcelBorderStyle.Thin;

            workSheet.Columns[2, 2].AutoFit();



            //Atenção que se escolhermos mais do que 24 anos isto vai dar erro

            for (int i = 0; i < Math.Min(numberOfYears, revenueValues.Count()); i++)
            {
                string aux = GetExcelColumnName(3 + i);
                string aux2 = "='P&L'!" + aux + "3";
                workSheet.Cells[3, 3 + i].Formula = "='P&L'!" + aux + "3";
                workSheet.Cells[3, 3 + i].Style.Numberformat.Format = DateTimeFormatInfo.CurrentInfo.ShortDatePattern;
                workSheet.Cells[3, 3 + i].Style.Font.Bold = true;
                workSheet.Cells[3, 3 + i].Style.Font.Color.SetColor(Cores.CorTexto);
                workSheet.Cells[3, 3 + i].Style.Border.Bottom.Style = ExcelBorderStyle.Thin;

                workSheet.Cells[2, 3 + i].Style.Font.Bold = true;
                workSheet.Cells[2, 3 + i].Style.Font.Color.SetColor(Color.White);

                workSheet.Cells[2, 3 + i].Style.Fill.PatternType = ExcelFillStyle.Solid;
                workSheet.Cells[2, 3 + i].Style.Fill.BackgroundColor.SetColor(Cores.CorPrincipal);
                workSheet.Columns[3 + i].Width = 14;

            }
            int row1 = 5;

            AnalysisConstruction construction = new AnalysisConstruction();


            //Cash and equivalents
            try
            {
                List<double> revenueValues1 = new List<double>();
                List<double> PropertyValues = new List<double>();

                foreach (FinancialStatements item in BalanceSheet)
                {

                    PropertyValues.Add(item.CashAndCashEquivalents);
                }

                foreach (FinancialStatements item in incomestatement)
                {
                    revenueValues1.Add(item.Revenue);
                }


                if (PropertyValues.Sum() != 0)
                {
                    construction.CommonAnalysisRefRevenueBS(workSheet, row1, numberOfYears, "Cash and cash equivalents", PropertyValues, revenueValues1, "N/A");
                    row1 = row1 + 7;
                }
            }
            catch (Exception er)
            {

                throw er;
            }




            //Short-term investments
            try
            {
                List<double> revenueValues1 = new List<double>();
                List<double> PropertyValues = new List<double>();

                foreach (FinancialStatements item in BalanceSheet)
                {

                    PropertyValues.Add(item.ShortTermInvestments);
                }
                foreach (FinancialStatements item in incomestatement)
                {
                    revenueValues1.Add(item.Revenue);
                }



                double teste1 = PropertyValues.Sum();
                if (teste1 != 0)
                {
                    construction.CommonAnalysisRefRevenueBS(workSheet, row1, numberOfYears, "Short term investments", PropertyValues, revenueValues1, "N/A");
                    row1 = row1 + 7;
                }
            }
            catch (Exception er)
            {

                throw er;
            }



            //Accounts receivable
            try
            {
                List<double> revenueValues1 = new List<double>();
                List<double> PropertyValues = new List<double>();

                foreach (FinancialStatements item in BalanceSheet)
                {

                    PropertyValues.Add(item.AccountsReceivables);
                }
                foreach (FinancialStatements item in incomestatement)
                {
                    revenueValues1.Add(item.Revenue);
                }

                if (PropertyValues.Sum() != 0)
                {
                    construction.CommonAnalysisRefRevenueBS(workSheet, row1, numberOfYears, "Accounts receivable, net", PropertyValues, revenueValues1,"N/A");
                    row1 = row1 + 7;
                }
            }
            catch (Exception er)
            {

                throw er;
            }




            //Inventory
            try
            {
                List<double> revenueValues1 = new List<double>();
                List<double> PropertyValues = new List<double>();

                foreach (FinancialStatements item in BalanceSheet)
                {

                    PropertyValues.Add(item.Inventory);
                }
                foreach (FinancialStatements item in incomestatement)
                {
                    revenueValues1.Add(item.Revenue);
                }

                if (PropertyValues.Sum() != 0)
                {
                    construction.CommonAnalysisRefRevenueBS(workSheet, row1, numberOfYears, "Inventory",PropertyValues, revenueValues1, "N/A");
                    row1 = row1 + 7;
                }
            }
            catch (Exception er)
            {

                throw er;
            }




            //Other current assets
            try
            {
                List<double> revenueValues1 = new List<double>();
                List<double> PropertyValues = new List<double>();

                foreach (FinancialStatements item in BalanceSheet)
                {

                    PropertyValues.Add(item.OtherCurrentAssets);
                }
                foreach (FinancialStatements item in incomestatement)
                {
                    revenueValues1.Add(item.Revenue);
                }

                if (PropertyValues.Sum() != 0)
                {
                    construction.CommonAnalysisRefRevenueBS(workSheet, row1, numberOfYears, "Other current assets", PropertyValues, revenueValues1, "N/A");
                    row1 = row1 + 7;
                }
            }
            catch (Exception er)
            {

                throw er;
            }



            //Property plant and equipment
            try
            {
                List<double> revenueValues1 = new List<double>();
                List<double> PropertyValues = new List<double>();

                foreach (FinancialStatements item in BalanceSheet)
                {

                    PropertyValues.Add(item.propertyPlantEquipmentNet);
                }
                foreach (FinancialStatements item in incomestatement)
                {
                    revenueValues1.Add(item.Revenue);
                }

                if (PropertyValues.Sum() != 0)
                {
                    construction.CommonAnalysisRefRevenueBS(workSheet, row1, numberOfYears, "Property plant and equipment", PropertyValues, revenueValues1, "N/A");
                    row1 = row1 + 7;
                }
            }
            catch (Exception er)
            {

                throw er;
            }


            //Goodwill
            try
            {
                List<double> revenueValues1 = new List<double>();
                List<double> PropertyValues = new List<double>();

                foreach (FinancialStatements item in BalanceSheet)
                {

                    PropertyValues.Add(item.Goodwill);
                }
                foreach (FinancialStatements item in incomestatement)
                {
                    revenueValues1.Add(item.Revenue);
                }

                if (PropertyValues.Sum() != 0)
                {
                    construction.CommonAnalysisRefRevenueBS(workSheet, row1, numberOfYears, "Goodwill", PropertyValues, revenueValues1, "N/A");
                    row1 = row1 + 7;
                }
            }
            catch (Exception er)
            {

                throw er;
            }




            //Intangible
            try
            {
                List<double> revenueValues1 = new List<double>();
                List<double> PropertyValues = new List<double>();

                foreach (FinancialStatements item in BalanceSheet)
                {

                    PropertyValues.Add(item.IntangibleAssets);
                }
                foreach (FinancialStatements item in incomestatement)
                {
                    revenueValues1.Add(item.Revenue);
                }

                if (PropertyValues.Sum() != 0)
                {
                    construction.CommonAnalysisRefRevenueBS(workSheet, row1, numberOfYears, "Intangible assets", PropertyValues, revenueValues1, "N/A");
                    row1 = row1 + 7;
                }
            }
            catch (Exception er)
            {

                throw er;
            }



            //Long-term investments
            try
            {
                List<double> revenueValues1 = new List<double>();
                List<double> PropertyValues = new List<double>();

                foreach (FinancialStatements item in BalanceSheet)
                {

                    PropertyValues.Add(item.LongtermInvestments);
                }
                foreach (FinancialStatements item in incomestatement)
                {
                    revenueValues1.Add(item.Revenue);
                }

                if (PropertyValues.Sum() != 0)
                {
                    construction.CommonAnalysisRefRevenueBS(workSheet, row1, numberOfYears, "Long term investments", PropertyValues, revenueValues1, "N/A");
                    row1 = row1 + 7;
                }
            }
            catch (Exception er)
            {

                throw er;
            }



            //Tax-assets
            try
            {
                List<double> revenueValues1 = new List<double>();
                List<double> PropertyValues = new List<double>();

                foreach (FinancialStatements item in BalanceSheet)
                {

                    PropertyValues.Add(item.TaxAssets);
                }
                foreach (FinancialStatements item in incomestatement)
                {
                    revenueValues1.Add(item.Revenue);
                }

                if (PropertyValues.Sum() != 0)
                {
                    construction.CommonAnalysisRefRevenueBS(workSheet, row1, numberOfYears, "Tax assets", PropertyValues, revenueValues1, "N/A");
                    row1 = row1 + 7;
                }
            }
            catch (Exception er)
            {

                throw er;
            }

            //Other non-current assets
            try
            {
                List<double> revenueValues1 = new List<double>();
                List<double> PropertyValues = new List<double>();

                foreach (FinancialStatements item in BalanceSheet)
                {

                    PropertyValues.Add(item.OtherNonCurrentAssets);
                }
                foreach (FinancialStatements item in incomestatement)
                {
                    revenueValues1.Add(item.Revenue);
                }

                if (PropertyValues.Sum() != 0)
                {
                    construction.CommonAnalysisRefRevenueBS(workSheet, row1, numberOfYears, "Other non-current assets", PropertyValues, revenueValues1, "N/A");
                    row1 = row1 + 7;
                }
            }
            catch (Exception er)
            {

                throw er;
            }



            //accounts payable
            try
            {
                List<double> revenueValues1 = new List<double>();
                List<double> PropertyValues = new List<double>();

                foreach (FinancialStatements item in BalanceSheet)
                {

                    PropertyValues.Add(item.AccountPayables);
                }
                foreach (FinancialStatements item in incomestatement)
                {
                    revenueValues1.Add(item.Revenue);
                }

                if (PropertyValues.Sum() != 0)
                {
                    construction.CommonAnalysisRefRevenueBS(workSheet, row1, numberOfYears, "Accounts payable", PropertyValues, revenueValues1, "N/A");
                    row1 = row1 + 7;
                }
            }
            catch (Exception er)
            {

                throw er;
            }





            //short-term debt
            try
            {
                List<double> revenueValues1 = new List<double>();
                List<double> PropertyValues = new List<double>();

                foreach (FinancialStatements item in BalanceSheet)
                {

                    PropertyValues.Add(item.ShortTermDebt);
                }
                foreach (FinancialStatements item in incomestatement)
                {
                    revenueValues1.Add(item.Revenue);
                }

                if (PropertyValues.Sum() != 0)
                {
                    construction.CommonAnalysisRefRevenueBS(workSheet, row1, numberOfYears, "Short term debt", PropertyValues, revenueValues1, "N/A");
                    row1 = row1 + 7;
                }
            }
            catch (Exception er)
            {

                throw er;
            }




            //tax payables
            try
            {
                List<double> revenueValues1 = new List<double>();
                List<double> PropertyValues = new List<double>();

                foreach (FinancialStatements item in BalanceSheet)
                {

                    PropertyValues.Add(item.TaxPayables);
                }
                foreach (FinancialStatements item in incomestatement)
                {
                    revenueValues1.Add(item.Revenue);
                }

                if (PropertyValues.Sum() != 0)
                {
                    construction.CommonAnalysisRefRevenueBS(workSheet, row1, numberOfYears, "Tax payables", PropertyValues, revenueValues1, "N/A");
                    row1 = row1 + 7;
                }
            }
            catch (Exception er)
            {

                throw er;
            }



            //Deferred revenue
            try
            {
                List<double> revenueValues1 = new List<double>();
                List<double> PropertyValues = new List<double>();

                foreach (FinancialStatements item in BalanceSheet)
                {

                    PropertyValues.Add(item.DeferredRevenue);
                }
                foreach (FinancialStatements item in incomestatement)
                {
                    revenueValues1.Add(item.Revenue);
                }

                if (PropertyValues.Sum() != 0)
                {
                    construction.CommonAnalysisRefRevenueBS(workSheet, row1, numberOfYears, "Deferred revenue", PropertyValues, revenueValues1, "N/A");
                    row1 = row1 + 7;
                }
            }
            catch (Exception er)
            {

                throw er;
            }




            //Other current liabilities
            try
            {
                List<double> revenueValues1 = new List<double>();
                List<double> PropertyValues = new List<double>();

                foreach (FinancialStatements item in BalanceSheet)
                {

                    PropertyValues.Add(item.OtherCurrentLiabilities);
                }
                foreach (FinancialStatements item in incomestatement)
                {
                    revenueValues1.Add(item.Revenue);
                }

                if (PropertyValues.Sum() != 0)
                {
                    construction.CommonAnalysisRefRevenueBS(workSheet, row1, numberOfYears, "Other current liabilities", PropertyValues, revenueValues1, "N/A");
                    row1 = row1 + 7;
                }
            }
            catch (Exception er)
            {

                throw er;
            }


            //Lon-term debt
            try
            {
                List<double> revenueValues1 = new List<double>();
                List<double> PropertyValues = new List<double>();

                foreach (FinancialStatements item in BalanceSheet)
                {

                    PropertyValues.Add(item.LongTermDebt);
                }
                foreach (FinancialStatements item in incomestatement)
                {
                    revenueValues1.Add(item.Revenue);
                }

                if (PropertyValues.Sum() != 0)
                {
                    construction.CommonAnalysisRefRevenueBS(workSheet, row1, numberOfYears, "Long term debt", PropertyValues, revenueValues1, "N/A");
                    row1 = row1 + 7;
                }
            }
            catch (Exception er)
            {

                throw er;
            }



            //deferred revenue, non-current
            try
            {
                List<double> revenueValues1 = new List<double>();
                List<double> PropertyValues = new List<double>();

                foreach (FinancialStatements item in BalanceSheet)
                {

                    PropertyValues.Add(item.DeferredRevenueNonCurrent);
                }
                foreach (FinancialStatements item in incomestatement)
                {
                    revenueValues1.Add(item.Revenue);
                }

                if (PropertyValues.Sum() != 0)
                {
                    construction.CommonAnalysisRefRevenueBS(workSheet, row1, numberOfYears, "Deferred revenue, non-current",  PropertyValues, revenueValues1, "N/A");
                    row1 = row1 + 7;
                }
            }
            catch (Exception er)
            {

                throw er;
            }




            //deferred tax liabilities, non-current
            try
            {
                List<double> revenueValues1 = new List<double>();
                List<double> PropertyValues = new List<double>();

                foreach (FinancialStatements item in BalanceSheet)
                {

                    PropertyValues.Add(item.DeferredTaxLiabilitiesNonCurrent);
                }
                foreach (FinancialStatements item in incomestatement)
                {
                    revenueValues1.Add(item.Revenue);
                }

                if (PropertyValues.Sum() != 0)
                {
                    construction.CommonAnalysisRefRevenueBS(workSheet, row1, numberOfYears, "Deferred tax liabilities, non-current",  PropertyValues, revenueValues1, "N/A");
                    row1 = row1 + 7;
                }
            }
            catch (Exception er)
            {

                throw er;
            }


            //other liabilities
            try
            {
                List<double> revenueValues1 = new List<double>();
                List<double> PropertyValues = new List<double>();

                foreach (FinancialStatements item in BalanceSheet)
                {

                    PropertyValues.Add(item.OtherLiabilities);
                }
                foreach (FinancialStatements item in incomestatement)
                {
                    revenueValues1.Add(item.Revenue);
                }

                if (PropertyValues.Sum() != 0)
                {
                    construction.CommonAnalysisRefRevenueBS(workSheet, row1, numberOfYears, "Other liabilities",  PropertyValues, revenueValues1, "N/A");
                    row1 = row1 + 7;
                }
            }
            catch (Exception er)
            {

                throw er;
            }




            //other non-current liabilities
            try
            {
                List<double> revenueValues1 = new List<double>();
                List<double> PropertyValues = new List<double>();

                foreach (FinancialStatements item in BalanceSheet)
                {

                    PropertyValues.Add(item.OtherNonCurrentLiabilities);
                }
                foreach (FinancialStatements item in incomestatement)
                {
                    revenueValues1.Add(item.Revenue);
                }

                if (PropertyValues.Sum() != 0)
                {
                    construction.CommonAnalysisRefRevenueBS(workSheet, row1, numberOfYears, "Other non-current liabilities", PropertyValues, revenueValues1,"N/A");
                    row1 = row1 + 7;
                }
            }
            catch (Exception er)
            {

                throw er;
            }
        }
    }
}
