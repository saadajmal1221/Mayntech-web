using Mayntech___Individual_Solution.Auxiliar.Auxiliar.FMP;
using Mayntech___Individual_Solution.Auxiliar.Auxiliar.Other;
using OfficeOpenXml;
using OfficeOpenXml.Style;
using System.Drawing;
using System.Globalization;

namespace Mayntech___Individual_Solution.Auxiliar.Auxiliar
{
    public class IncomeStatementConstruction
    {
        public async Task CreatePL(ExcelPackage package, List<FinancialStatements> incomeStatement, int col, int row, 
            string companyName, PeersProfile companyProfile)
        {
            var workSheet = package.Workbook.Worksheets.Add("P&L");
            workSheet.View.ShowGridLines = false;
            workSheet.View.FreezePanes(4, 3);

            int quarters = Quarters(companyProfile);

            ExcelNextCol columnName = new ExcelNextCol();

            // Cria a coluna azul em cima da tabela
            workSheet.Cells[2, 2].Value = "P&L - " + companyName;
            workSheet.Cells[2, 2].Style.Font.Bold = true;
            workSheet.Cells[2, 2].Style.Font.Color.SetColor(Color.White);

            // atenção que este "H2" está hardcoded. Tem de ser refeito para variar com o número de anos
            workSheet.Cells["B2"].Style.Fill.PatternType = ExcelFillStyle.Solid;
            workSheet.Cells["B2"].Style.Fill.BackgroundColor.SetColor(Cores.CorPrincipal);

            //criar um field com o valor ser em milhoes ou thousands


            incomeStatement.Reverse();

            foreach (FinancialStatements item in incomeStatement)
            {
                int row1 = 5;
                ConstructionFinancialStatementsSupport support = new();

                workSheet.Cells[2, 3 + col].Style.Fill.PatternType = ExcelFillStyle.Solid;
                workSheet.Cells[2, 3 + col].Style.Fill.BackgroundColor.SetColor(Cores.CorPrincipal);

                // cria o header
                workSheet.Cells[3, 2].Value = "Description (in '000 " + item.ReportedCurrency + ")";
                workSheet.Cells[3, 2].Style.Font.Bold = true;
                workSheet.Cells[3, 2].Style.Font.Color.SetColor(Cores.CorTexto);
                workSheet.Cells[3, 2].Style.Border.Bottom.Style = ExcelBorderStyle.Thin;

                // inserir a data com as formatações
                workSheet.Cells[row1 - 2, 3 + col].Value = item.Date;
                workSheet.Cells[row1 - 2, 3 + col].Style.Numberformat.Format = DateTimeFormatInfo.CurrentInfo.ShortDatePattern;
                workSheet.Cells[row1 - 2, 3 + col].Style.Font.Bold = true;
                workSheet.Cells[row1 - 2, 3 + col].Style.Font.Color.SetColor(Cores.CorTexto);
                workSheet.Cells[row1 - 2, 3 + col].Style.Border.Bottom.Style = ExcelBorderStyle.Thin;

                

                // define o height da linha entre o cabeçalho e os valores
                workSheet.Row(row1 - 1).Height = 4;

                if (item.Revenue != null)
                {
                    //workSheet.Cells[row1, 2].Value = "Revenue";
                    //workSheet.Cells[row1, 3 + col].Value = item.Revenue / 1000;
                    //workSheet.Cells[row1, 3+col].Style.Numberformat.Format = "#,##0 ;(#,##0)";
                    //row1++;

                    support.CommonCaption("Revenue", item.Revenue, col, row1, workSheet, item);
                    row1++;
                }


                if (item.CostOfRevenue != null)
                {
                    support.CommonCaptionNegative("Cost of Revenue", item.CostOfRevenue, col, row1, workSheet, item);

                    row1++;
                }

                if (item.GrossProfit != null)
                {
                    //support.CaptionTotal("Gross Profit", item.GrossProfit, col, row1, workSheet, item);

                    string colAux = columnName.GetExcelColumnName(col + 3);
                    workSheet.Cells[row1, 2].Value = "Gross Profit";
                    workSheet.Cells[row1, 2].Style.Font.Bold = true;
                    workSheet.Cells[row1, col + 3].Formula = "=" + colAux + "5+" + colAux + "6";
                    workSheet.Cells[row1, col + 3].Style.Numberformat.Format = "#,##0;(#,##0);-";
                    workSheet.Cells[row1, col + 3].Style.Font.Bold = true;

                    row1++;
                }

                if (item.GrossProfitRatio != null)
                {
                    //support.CaptionRatio("Gross Profit Ratio", item.GrossProfitRatio, col, row1, workSheet, item);
                    string colAux = columnName.GetExcelColumnName(col+3);

                    workSheet.Cells[row1, 2].Value = "Gross Profit Ratio";
                    workSheet.Cells[row1, 2].Style.Font.Color.SetColor(Color.Gray);
                    workSheet.Cells[row1, col+3].Formula = "=" + colAux + "7/" + colAux + "5";
                    workSheet.Cells[row1, col + 3].Style.Numberformat.Format = "0.00%";
                    workSheet.Cells[row1, col + 3].Style.Font.Color.SetColor(Color.Gray);


                    row1++;
                    //duas vezes para criar um espaço
                    row1++;

                }
                if (item.OperatingExpenses != null)
                {
                    
                    string colAux = columnName.GetExcelColumnName(col + 3);
                    workSheet.Cells[row1, 2].Value = "Operating Expenses";
                    workSheet.Cells[row1, col + 3].Formula = "=" + colAux + "11+" + colAux + "12+" + colAux + "13";
                    workSheet.Cells[row1, col + 3].Style.Numberformat.Format = "#,##0;(#,##0);-";

                    

                }
                row1++;
                if (item.ResearchAndDevelopmentExpenses != null)
                {
                    support.CommonSubCaption("Research and development Expenses", item.ResearchAndDevelopmentExpenses, col, row1, workSheet, item);                 
                }
                row1++;
                //if (item.GeneralAndAdministrativeExpenses != null)
                //{
                //    support.CommonSubCaption("General and Administrative Expenses", item.GeneralAndAdministrativeExpenses, col, row1, workSheet, item, cor);

                //    row1++;


                //}
                if (item.SellingGeneralAndAdministrativeExpenses != null)
                {
                    support.CommonSubCaption("Selling General and Administrative Expenses", item.SellingGeneralAndAdministrativeExpenses, col, row1, workSheet, item);

                }
                row1++;


                if (item.otherExpenses != null)
                {
                    support.CommonSubCaption("Other operating Income/Expenses", item.otherExpenses, col, row1, workSheet, item);

                }
                row1++;
                if (item.OperatingExpenses != null)
                {
                    support.CommonCaptionNegative("Other", -(item.OperatingIncome - (item.GrossProfit - item.OperatingExpenses)), col, row1, workSheet, item);

                }
                row1++;
                if (item.OperatingIncome != null)
                {
                    string colAux = columnName.GetExcelColumnName(col + 3);
                    workSheet.Cells[row1, 2].Value = "Operating Income";
                    workSheet.Cells[row1, 2].Style.Font.Bold = true;
                    workSheet.Cells[row1, col + 3].Formula = "=" + colAux + "7+" + colAux + "10+" + colAux + "14";
                    workSheet.Cells[row1, col + 3].Style.Numberformat.Format = "#,##0;(#,##0);-";
                    workSheet.Cells[row1, col + 3].Style.Font.Bold = true;

                }
                row1++;
                if (item.OperatingIncomeRatio != null)
                {
                    //support.CaptionRatio("Operating Income Ratio", item.OperatingIncomeRatio, col, row1, workSheet, item);
                    string colAux = columnName.GetExcelColumnName(col + 3);

                    workSheet.Cells[row1, 2].Value = "Operating Income Ratio";
                    workSheet.Cells[row1, 2].Style.Font.Color.SetColor(Color.Gray);
                    workSheet.Cells[row1, col + 3].Formula = "=" + colAux + "15/" + colAux + "5";
                    workSheet.Cells[row1, col + 3].Style.Numberformat.Format = "0.00%";
                    workSheet.Cells[row1, col + 3].Style.Font.Color.SetColor(Color.Gray);

                }
                row1++;
                row1++;
                if (item.TotalOtherIncomeExpensesNet != null)
                {
                    support.CommonCaption("Total other income/expenses, net", item.TotalOtherIncomeExpensesNet, col, row1, workSheet, item);
                }
                row1++;

                if (item.IncomeBeforeTax != null)
                {
                    //support.CaptionTotal("Income Before Tax", (double)item.IncomeBeforeTax, col, row1, workSheet, item);

                    string colAux = columnName.GetExcelColumnName(col + 3);
                    workSheet.Cells[row1, 2].Value = "Income Before Tax";
                    workSheet.Cells[row1, 2].Style.Font.Bold = true;
                    workSheet.Cells[row1, col + 3].Formula = "=" + colAux + "15+" + colAux + "18";
                    workSheet.Cells[row1, col + 3].Style.Numberformat.Format = "#,##0;(#,##0);-";
                    workSheet.Cells[row1, col + 3].Style.Font.Bold = true;

                }
                row1++;
                if (item.IncomeBeforeTaxRatio != null)
                {
                    //support.CaptionRatio("IncomeBeforeTaxRatio", item.IncomeBeforeTaxRatio, col, row1, workSheet, item);

                    string colAux = columnName.GetExcelColumnName(col + 3);

                    workSheet.Cells[row1, 2].Value = "Income Before Tax Ratio";
                    workSheet.Cells[row1, 2].Style.Font.Color.SetColor(Color.Gray);
                    workSheet.Cells[row1, col + 3].Formula = "=" + colAux + "19/" + colAux + "5";
                    workSheet.Cells[row1, col + 3].Style.Numberformat.Format = "0.00%";
                    workSheet.Cells[row1, col + 3].Style.Font.Color.SetColor(Color.Gray);

                }
                row1++;
                row1++;


                if (item.IncomeTaxExpense != null)
                {
                    support.CommonCaptionNegative("Income Tax Expense", (double)item.IncomeTaxExpense, col, row1, workSheet, item);
                   

                }
                row1++;
                if (item.IncomeBeforeTax != null)
                {
                    //support.CaptionTotal("Consolidated Net Income", (double)item.IncomeBeforeTax - (double)item.IncomeTaxExpense, col, row1, workSheet, item);
                    string colAux = columnName.GetExcelColumnName(col + 3);
                    workSheet.Cells[row1, 2].Value = "Consolidated Net Income";
                    workSheet.Cells[row1, 2].Style.Font.Bold = true;
                    workSheet.Cells[row1, col + 3].Formula = "=" + colAux + "19+" + colAux + "22";
                    workSheet.Cells[row1, col + 3].Style.Numberformat.Format = "#,##0;(#,##0);-";
                    workSheet.Cells[row1, col + 3].Style.Font.Bold = true;

                }
                row1++;
                row1++;
                if (item.IncomeBeforeTax != null)
                {
                    support.CommonCaptionNegative("Other charges (including minority interests)", ((double)item.IncomeBeforeTax - (double)item.IncomeTaxExpense - item.NetIncome) , col, row1, workSheet, item);
                    

                }
                row1++;
                if (item.NetIncome != null)
                {
                    //support.CaptionTotal("Net Income attributable to the group", item.NetIncome, col, row1, workSheet, item);
                    string colAux = columnName.GetExcelColumnName(col + 3);
                    workSheet.Cells[row1, 2].Value = "Net Income attributable to the group";
                    workSheet.Cells[row1, 2].Style.Font.Bold = true;
                    workSheet.Cells[row1, col + 3].Formula = "=" + colAux + "23+" + colAux + "25";
                    workSheet.Cells[row1, col + 3].Style.Numberformat.Format = "#,##0;(#,##0);-";
                    workSheet.Cells[row1, col + 3].Style.Font.Bold = true;

                }
                row1++;
                if (item.NetIncomeRatio != null)
                {
                    support.CaptionRatio("Net Income Ratio", item.NetIncomeRatio, col, row1, workSheet, item);

                    string colAux = columnName.GetExcelColumnName(col + 3);
                    workSheet.Cells[row1, 2].Value = "Net Income Ratio";
                    workSheet.Cells[row1,2].Style.Font.Color.SetColor(Color.Gray);
                    workSheet.Cells[row1, col + 3].Formula = "=" + colAux + "26/" + colAux + "5";
                    workSheet.Cells[row1, col + 3].Style.Numberformat.Format = "0.00%";
                    workSheet.Cells[row1, col + 3].Style.Font.Color.SetColor(Color.Gray);

                    //adicionar a border line
                }
                row1++;
                row1++;

                workSheet.Cells[row1 + 5, 2].Value = "Other Captions";
                workSheet.Cells[row1 + 5, 2].Style.Font.Bold = true;

                row1++;

                if (item.EBITDA != null)
                {
                    support.CommonCaption("EBITDA", item.EBITDA, col, row1, workSheet, item);
                    
                }
                row1++;
                if (item.EBITDARatio != null)
                {
                    //support.CaptionRatio("EBITDA ratio", item.EBITDARatio, col, row1, workSheet, item);

                    string colAux = columnName.GetExcelColumnName(col + 3);
                    workSheet.Cells[row1, 2].Value = "EBITDA ratio";
                    workSheet.Cells[row1, 2].Style.Font.Color.SetColor(Color.Gray);
                    workSheet.Cells[row1, col + 3].Formula = "=" + colAux + "30/" + colAux + "5";
                    workSheet.Cells[row1, col + 3].Style.Numberformat.Format = "0.00%";
                    workSheet.Cells[row1, col + 3].Style.Font.Color.SetColor(Color.Gray);

                }
                row1++;
                if (item.DepreciationAndAmortization != null)
                {
                    support.CommonCaptionNegative("Depreciation and amortization", (double)item.DepreciationAndAmortization, col, row1, workSheet, item);
                    
                }
                row1++;
                if (item.interestIncome != null)
                {
                    support.CommonCaption("Interest Income", (double)item.interestIncome, col, row1, workSheet, item);
                    
                }
                row1++;
                if (item.InterestExpense != null)
                {
                    support.CommonCaptionNegative("Interest Expense", (double)item.InterestExpense, col, row1, workSheet, item);

                }
                row1++;
                row1++;

                if (item.EPS != null)
                {
                    workSheet.Cells[row1, 2].Value = "EPS";
                    workSheet.Cells[row1, 3 + col].Value = item.EPS;
                    workSheet.Cells[row1, 3 + col].Style.Numberformat.Format = "#,##0.00;(#,##0.00);-";
                    row1++;
                }

                if (item.EPSDiluted != null)
                {
                    workSheet.Cells[row1, 2].Value = "EPS diluted";
                    workSheet.Cells[row1, 3 + col].Value = item.EPSDiluted;
                    workSheet.Cells[row1, 3 + col].Style.Numberformat.Format = "#,##0.00;(#,##0.00);-";
                    
                }
                row1++;
                if (item.weightedAverageShsOut != null)
                {
                    workSheet.Cells[row1, 2].Value = "Weighted average shares outstanding";
                    workSheet.Cells[row1, 3 + col].Value = item.weightedAverageShsOut;
                    workSheet.Cells[row1, 3 + col].Style.Numberformat.Format = "#,##0;(#,##0);-";
                    
                }
                row1++;
                if (item.weightedAverageShsOutDil != null)
                {
                    workSheet.Cells[row1, 2].Value = "Weighted average shares outstanding diluted";
                    workSheet.Cells[row1, 3 + col].Value = item.weightedAverageShsOutDil;
                    workSheet.Cells[row1, 3 + col].Style.Numberformat.Format = "#,##0;(#,##0);-";
                    workSheet.Cells[row1, 3 + col].Style.Border.Bottom.Style = ExcelBorderStyle.Thin;
                    workSheet.Cells[row1, 2].Style.Border.Bottom.Style = ExcelBorderStyle.Thin;
                    
                }
                row1++;



                workSheet.Columns[2, 3 + col].AutoFit();
                col++;

            }

            if (quarters > 0)
            {  

                for (int i = quarters - 1; i > -1; i--)
                {
                    if (i == 0)
                    {
                        workSheet.Columns[3 + incomeStatement.Count()].Width = 3;
                    }
                    int column = 3 + (quarters - i) + incomeStatement.Count();
                    int row1 = 5;

                    if (i == quarters-1)
                    {
                        workSheet.Cells[2, column].Value = "Last Year Quarters ";
                        workSheet.Cells[2, column].Style.Font.Bold = true;
                        workSheet.Cells[2, column].Style.Font.Color.SetColor(Color.White);
                    }
                    workSheet.Cells[2, column].Style.Fill.PatternType = ExcelFillStyle.Solid;
                    workSheet.Cells[2, column].Style.Fill.BackgroundColor.SetColor(Cores.CorPrincipal);


                    string columnAux = columnName.GetExcelColumnName(column);


                    for (int a = 0; a < 36; a++)
                    {
                        workSheet.Cells[4 + a, column].Style.Fill.PatternType = ExcelFillStyle.Solid;
                        workSheet.Cells[4 + a, column].Style.Fill.BackgroundColor.SetColor(Cores.CorCinzentoBackground);

                    }

                    ConstructionFinancialStatementsSupport support = new();

                    workSheet.Cells[row1 - 2, column].Value = companyProfile.financialsQuarter.income[i].Date;
                    workSheet.Cells[row1 - 2, column].Style.Numberformat.Format = DateTimeFormatInfo.CurrentInfo.ShortDatePattern;
                    workSheet.Cells[row1 - 2, column].Style.Font.Bold = true;
                    workSheet.Cells[row1 - 2, column].Style.Font.Color.SetColor(Cores.CorTexto);
                    workSheet.Cells[row1 - 2, column].Style.Border.Bottom.Style = ExcelBorderStyle.Thin;

                    //Revenue
                    workSheet.Cells[row1, column].Value = companyProfile.financialsQuarter.income[i].Revenue / 1000;
                    workSheet.Cells[row1, column].Style.Numberformat.Format = "#,##0;(#,##0);-";
                    row1++;

                    //COGS
                    workSheet.Cells[row1, column].Value = -companyProfile.financialsQuarter.income[i].CostOfRevenue / 1000;
                    workSheet.Cells[row1, column].Style.Numberformat.Format = "#,##0;(#,##0);-";
                    row1++;


                    //Gross Profit
                    workSheet.Cells[row1, column].Formula = "=" + columnAux + "5+" + columnAux + "6";
                    workSheet.Cells[row1, column].Style.Numberformat.Format = "#,##0;(#,##0);-";
                    workSheet.Cells[row1, column].Style.Font.Bold = true;
                    row1++;

                    //Gross Profit Ratio
                    workSheet.Cells[row1, column].Formula = "=" + columnAux + "7/" + columnAux + "5";
                    
                    workSheet.Cells[row1, column].Style.Numberformat.Format = "0.00%";
                    workSheet.Cells[row1, column].Style.Font.Color.SetColor(Color.Gray);
                    row1 += 2;

                    //Operating Expenses
                    workSheet.Cells[row1, column].Formula = "=" + columnAux + "11+" + columnAux + "12+" + columnAux + "13";
                    
                    workSheet.Cells[row1, column].Style.Numberformat.Format = "#,##0;(#,##0);-";
                    row1++;

                    //Research and development Expenses
                    workSheet.Cells[row1, column].Value = -companyProfile.financialsQuarter.income[i].ResearchAndDevelopmentExpenses / 1000;
                    workSheet.Cells[row1, column].Style.Numberformat.Format = "#,##0;(#,##0);-";
                    row1++;

                    //Selling And Administrative
                    workSheet.Cells[row1, column].Value = -companyProfile.financialsQuarter.income[i].SellingGeneralAndAdministrativeExpenses / 1000;
                    workSheet.Cells[row1, column].Style.Numberformat.Format = "#,##0;(#,##0);-";
                    row1++;

                    //Other Operating Expenses
                    workSheet.Cells[row1, column].Value = -companyProfile.financialsQuarter.income[i].otherExpenses / 1000;
                    workSheet.Cells[row1, column].Style.Numberformat.Format = "#,##0;(#,##0);-";
                    row1++;

                    //Other
                    workSheet.Cells[row1, column].Value = -(companyProfile.financialsQuarter.income[i].OperatingIncome - (companyProfile.financialsQuarter.income[i].GrossProfit - companyProfile.financialsQuarter.income[i].OperatingExpenses))/1000;
                    workSheet.Cells[row1, column].Style.Numberformat.Format = "0.00%";
                    row1++;

                    //Operating Income
                    workSheet.Cells[row1, column].Formula = "=" + columnAux + "7+" + columnAux + "10+" + columnAux + "14";
                    workSheet.Cells[row1, column].Style.Numberformat.Format = "#,##0;(#,##0);-";
                    workSheet.Cells[row1, column].Style.Font.Bold = true;
                    row1++;

                    //Operating Income Ratio
                    workSheet.Cells[row1, column].Formula = "=" + columnAux + "15/" + columnAux + "5";
                    workSheet.Cells[row1, column].Style.Numberformat.Format = "0.00%";
                    workSheet.Cells[row1, column].Style.Font.Color.SetColor(Color.Gray);
                    row1 += 2;

                    //Total other income/expenses
                    workSheet.Cells[row1, column].Value = companyProfile.financialsQuarter.income[i].TotalOtherIncomeExpensesNet / 1000;
                    workSheet.Cells[row1, column].Style.Numberformat.Format = "#,##0;(#,##0);-";
                    row1++;

                    //income before tax
                    workSheet.Cells[row1, column].Formula = "=" + columnAux + "15+" + columnAux + "18";
                    workSheet.Cells[row1, column].Style.Numberformat.Format = "#,##0;(#,##0);-";
                    workSheet.Cells[row1, column].Style.Font.Bold = true;
                    row1++;

                    //income before tax ratio
                    workSheet.Cells[row1, column].Formula = "=" + columnAux + "19/" + columnAux + "5";
                    workSheet.Cells[row1, column].Style.Numberformat.Format = "0.00%";
                    workSheet.Cells[row1, column].Style.Font.Color.SetColor(Color.Gray);
                    row1 += 2;

                    //income tax expense
                    workSheet.Cells[row1, column].Value = -companyProfile.financialsQuarter.income[i].IncomeTaxExpense / 1000;
                    workSheet.Cells[row1, column].Style.Numberformat.Format = "#,##0;(#,##0);-";
                    row1++;

                    //consolidated net income
                    workSheet.Cells[row1, column].Formula = "=" + columnAux + "19+" + columnAux + "22";
                    workSheet.Cells[row1, column].Style.Numberformat.Format = "#,##0;(#,##0);-";
                    workSheet.Cells[row1, column].Style.Font.Bold = true;
                    row1 += 2;

                    //Other Charges
                    workSheet.Cells[row1, column].Value = (companyProfile.financialsQuarter.income[i].NetIncome - (companyProfile.financialsQuarter.income[i].IncomeBeforeTax - companyProfile.financialsQuarter.income[i].IncomeTaxExpense)) / 1000;
                    workSheet.Cells[row1, column].Style.Numberformat.Format = "#,##0;(#,##0);-";

                    row1++;


                    //Net income attributable to the group
                    workSheet.Cells[row1, column].Formula = "=" + columnAux + "23+" + columnAux + "25";
                    workSheet.Cells[row1, column].Style.Numberformat.Format = "#,##0;(#,##0);-";
                    workSheet.Cells[row1, column].Style.Font.Bold = true;
                    row1++;

                    //Net income ratio
                    workSheet.Cells[row1, column].Formula = "=" + columnAux + "26/" + columnAux + "5";
                    workSheet.Cells[row1, column].Style.Numberformat.Format = "0.00%";
                    workSheet.Cells[row1, column].Style.Font.Color.SetColor(Color.Gray);
                    row1 += 3;

                    //EBITDA
                    workSheet.Cells[row1, column].Value = companyProfile.financialsQuarter.income[i].EBITDA / 1000;
                    workSheet.Cells[row1, column].Style.Numberformat.Format = "#,##0;(#,##0);-";
                    row1++;

                    //EBITDA Ratio
                    workSheet.Cells[row1, column].Formula = "=" + columnAux + "30/" + columnAux + "5";
                    workSheet.Cells[row1, column].Style.Numberformat.Format = "0.00%";
                    workSheet.Cells[row1, column].Style.Font.Color.SetColor(Color.Gray);
                    row1++;

                    //D&A
                    workSheet.Cells[row1, column].Value = -companyProfile.financialsQuarter.income[i].DepreciationAndAmortization / 1000;
                    workSheet.Cells[row1, column].Style.Numberformat.Format = "#,##0;(#,##0);-";
                    row1++;

                    //interestincome
                    workSheet.Cells[row1, column].Value = companyProfile.financialsQuarter.income[i].interestIncome / 1000;
                    workSheet.Cells[row1, column].Style.Numberformat.Format = "#,##0;(#,##0);-";
                    row1++;

                    //interestexpense
                    workSheet.Cells[row1, column].Value = companyProfile.financialsQuarter.income[i].InterestExpense / 1000;
                    workSheet.Cells[row1, column].Style.Numberformat.Format = "#,##0;(#,##0);-";
                    row1++;

                    workSheet.Columns[column].AutoFit();

                }
            }
            //workSheet.Cells["B:X"].AutoFitColumns();
        }

        public int Quarters(PeersProfile companyProfile)
        {
            DateTime AnnualReport = companyProfile.financialsAnnual.income[0].Date;

            int quarters = QuartersAux(AnnualReport, companyProfile, 0);

            return quarters;

        }

        public int QuartersAux(DateTime AnnualReport, PeersProfile companyProfile, int aux)
        {
            int output = 0;
            if (companyProfile.financialsQuarter.income[aux].Date > AnnualReport)
            {
                output += 1;
                output += QuartersAux(AnnualReport, companyProfile, aux + 1);

            }

            return output;
        }

    }
}
