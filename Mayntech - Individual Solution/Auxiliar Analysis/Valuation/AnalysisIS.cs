using OfficeOpenXml.Style;
using OfficeOpenXml;
using System.Drawing;
using System.Globalization;
using Microsoft.VisualBasic;
using Mayntech___Individual_Solution.Auxiliar.Analysis.Analysis_Support;
using Mayntech___Individual_Solution.Auxiliar.Auxiliar.FMP;
using Mayntech___Individual_Solution.Auxiliar.Auxiliar.Other;
using Mayntech___Individual_Solution.Pages;
using System.Data.Common;

namespace Mayntech___Individual_Solution.Auxiliar.Analysis
{
    public class AnalysisIS : ExcelNextCol
    {


        public async Task CreateISAnalysis(ExcelPackage package, List<FinancialStatements> incomeStatement, int col, int row, string companyName, int numberOfYears)
        {
            var workSheet = package.Workbook.Worksheets.Add("P&L - Analysis");
            workSheet.View.ShowGridLines = false;
            workSheet.View.ZoomScale = 80;

            List<double> revenueValues = new List<double>();
            foreach (FinancialStatements item in incomeStatement)
            {
                revenueValues.Add(item.Revenue);
            }

            // Cria a coluna azul em cima da tabela
            workSheet.Cells[2, 2].Value = "P&L Analysis - " + companyName;
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
            RevenueAnalysis revenue = new RevenueAnalysis();

            revenue.RevenueAnalysisConstruction(workSheet, row1, numberOfYears, revenueValues);
            row1 = row1 + 5;

            //COGS
            if (true)
            {
                List<double> revenueValues1 = new List<double>();
                List<double> PropertyValues = new List<double>();

                foreach (FinancialStatements item in incomeStatement)
                {
                    revenueValues1.Add(item.Revenue);
                    PropertyValues.Add(item.CostOfRevenue);
                }

                if (PropertyValues.Sum() != 0)
                {
                    construction.CommonAnalysisRefRevenue(workSheet, row1, numberOfYears, "Cost Of Revenue", PropertyValues, revenueValues1, "Negative");
                    row1 = row1 + 7;
                }
            }

            //R&D
            try
            {
                List<double> PropertyValues = new List<double>();
                List<double> revenueValues1 = new List<double>();
                foreach (FinancialStatements item in incomeStatement)
                {
                    revenueValues1.Add(item.Revenue);
                    PropertyValues.Add(item.ResearchAndDevelopmentExpenses);
                }
                if (PropertyValues.Sum() != 0)
                {

                    construction.CommonAnalysisRefRevenue(workSheet, row1, numberOfYears, "Research and Development expenses", PropertyValues, revenueValues1, "Negative");

                          
                    //Fim do teste



                    row1 = row1 + 7;
                }
            }
            catch (Exception er)
            {

                throw er;
            }



            //Selling and adminsitrative expenses
            try
            {

                List<double> revenueValues1 = new List<double>();
                List<double> PropertyValues = new List<double>();
                foreach (FinancialStatements item in incomeStatement)
                {
                    revenueValues1.Add(item.Revenue);
                    PropertyValues.Add(item.SellingGeneralAndAdministrativeExpenses);
                }
                if (PropertyValues.Sum() != 0)
                {
                    construction.CommonAnalysisRefRevenue(workSheet, row1, numberOfYears, "Selling general and administrative expenses", PropertyValues, revenueValues1, "Negative");
                    row1 = row1 + 7;
                }


            }
            catch (Exception)
            {

                throw;
            }



            //total other income/expenses Net
            try
            {

                List<double> revenueValues1 = new List<double>();
                List<double> PropertyValues = new List<double>();
                foreach (FinancialStatements item in incomeStatement)
                {
                    revenueValues1.Add(item.Revenue);
                    PropertyValues.Add(item.TotalOtherIncomeExpensesNet);
                }
                if (PropertyValues.Sum() != 0)
                {
                    construction.CommonAnalysisRefRevenue(workSheet, row1, numberOfYears, "Total other income/expenses, net", PropertyValues, revenueValues1, "N/A");
                    row1 = row1 + 7;
                }


            }
            catch (Exception)
            {

                throw;
            }


            //Income Tax expense
            try
            {

                List<double> revenueValues1 = new List<double>();
                List<double> PropertyValues = new List<double>();
                foreach (FinancialStatements item in incomeStatement)
                {
                    revenueValues1.Add(item.Revenue);
                    PropertyValues.Add((double)item.IncomeTaxExpense);
                }
                if (PropertyValues.Sum() != 0)
                {
                    construction.CommonAnalysisRefRevenue(workSheet, row1, numberOfYears, "Income Tax Expense", PropertyValues, revenueValues1, "Negative");
                    row1 = row1 + 7;
                }


            }
            catch (Exception)
            {

                throw;
            }

            
            string EndColumn = GetExcelColumnName(col +2 +numberOfYears);
            int auxRow = row1 + 3;
            int auxRow2 = auxRow + 1;
            int auxRowMinusOne = auxRow -1;
            workSheet.Cells["B" + auxRow + ":" + EndColumn + auxRow2].Merge = true;
            workSheet.Cells["B" + auxRow + ":" + EndColumn + auxRow2].Style.VerticalAlignment = ExcelVerticalAlignment.Top;
            workSheet.Cells["B" + auxRow + ":" + EndColumn + auxRow2].Style.HorizontalAlignment = ExcelHorizontalAlignment.Left;
            workSheet.Cells["B" + auxRow + ":" + EndColumn + auxRow2].Style.Fill.PatternType = ExcelFillStyle.Solid;
            workSheet.Cells["B" + auxRow + ":" + EndColumn + auxRow2].Style.Fill.BackgroundColor.SetColor(Color.LightGray);

            workSheet.Cells["B" + auxRow + ":" + EndColumn + auxRow2].Style.WrapText = true;

            workSheet.Cells["B" + auxRowMinusOne].Value = "Note:";
            workSheet.Cells["B" + auxRowMinusOne].Style.Font.Bold = true;

            workSheet.Cells["B" + auxRow].Value = "All estimations are for the next 3-5 years. All estimations are based purely on mathematical analysis and may not be representative of the future. ";
            

            string StartColumnWarning = GetExcelColumnName(col + 4+numberOfYears);
            string EndColumnWarning = GetExcelColumnName(col + 7 + numberOfYears);


            workSheet.Cells[StartColumnWarning +"2:" + EndColumnWarning +"2"].Merge = true;
            workSheet.Cells[StartColumnWarning + "2:" + EndColumnWarning + "2"].Style.VerticalAlignment = ExcelVerticalAlignment.Top;
            workSheet.Cells[StartColumnWarning + "2:" + EndColumnWarning + "2"].Style.HorizontalAlignment = ExcelHorizontalAlignment.Left;
            workSheet.Cells[StartColumnWarning + "2:" + EndColumnWarning + "2"].Style.Fill.PatternType = ExcelFillStyle.Solid;
            workSheet.Cells[StartColumnWarning + "2:" + EndColumnWarning + "2"].Style.Fill.BackgroundColor.SetColor(Color.LightGray);
            workSheet.Cells[StartColumnWarning + "2:" + EndColumnWarning + "2"].Style.WrapText = true;
            workSheet.Cells[StartColumnWarning + "2:" + EndColumnWarning + "2"].Value = "Please check the note in cell: B" + auxRow;  

        }
    }
}
