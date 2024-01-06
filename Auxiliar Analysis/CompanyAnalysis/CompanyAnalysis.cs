using Mayntech___Individual_Solution.Auxiliar.Auxiliar.FMP;
using Mayntech___Individual_Solution.Auxiliar.Auxiliar.Other;
using Mayntech___Individual_Solution.Auxiliar_Valuation.Financials_projections;
using Mayntech___Individual_Solution.Pages;
using OfficeOpenXml;
using OfficeOpenXml.Sparkline;
using OfficeOpenXml.Style;
using System.Data.Common;
using System.Drawing;
using System.Runtime.Intrinsics.X86;

namespace Mayntech___Individual_Solution.Auxiliar.CompanyAnalysis
{
    public class CompanyAnalysis : CompanyAnalysisCommentBuilder
    {
        public void financialSummary(ExcelPackage package, int numberOfYears, List<FinancialStatements> incomestatement,
            List<FinancialStatements> balances, List<FinancialStatements> cashFlow, Taxes tax)
        {
            ExcelNextCol nextCol = new ExcelNextCol();
            var worksheet = package.Workbook.Worksheets.Add("Analysis");
            worksheet.View.ShowGridLines = false;
            worksheet.View.ZoomScale = 80;
            worksheet.View.FreezePanes(4, 3);
            int row = 2;
            int col = 2;

            int numberOfColumns = Math.Min(5, numberOfYears);



            List<int> years = new List<int>();
            List<double> Revenue = new List<double>();
            
            List<double> GrossProfit = new List<double>();
            List<double> GrossMargin = new List<double>();
            
            List<double> OperatingMargin = new List<double>();
            List<double> OperatingProfit = new List<double>();

            List<double> Noplat = new List<double>();


            List<double> Cash = new List<double>();
            List<double> PPE = new List<double>();

            List<double> WorkingCapital = new List<double>();
            List<double> InvestedCapital = new List<double>();
            List<double> Goodwill = new List<double>();
            
            List<double> totalDebt = new List<double>();
            List<double> RetainedEarnings = new List<double>();
            List<double> CashFlowFromOperations = new List<double>();
            List<double> CashFlowFromFinancing = new List<double>();
            List<double> CashFlowFromInvesting = new List<double>();
            List<double> FreeCashFlow = new List<double>();

            List<FinancialStatements> balanceSheet = new List<FinancialStatements>();
            balanceSheet = SolutionModel.balances;
            balanceSheet.Reverse();
            int LastYear = int.Parse(balanceSheet[0].CalendarYear);

            List<double> auxWcOp = new List<double>();

            int OpWCAux = 0;
            for (int i = 0; i < balanceSheet.Count(); i++)
            {
                auxWcOp.Add((double)balanceSheet[i].NetReceivables + (double)balanceSheet[i].Inventory + (double)balanceSheet[i].OtherCurrentAssets - ((double)balanceSheet[i].AccountPayables + (double)balanceSheet[i].DeferredRevenue + balanceSheet[i].OtherCurrentLiabilities));
            }

            if (auxWcOp.Min()<0)
            {
                OpWCAux = 1;
            }

            List<FinancialStatements> incomeStatement = new List<FinancialStatements>();
            incomeStatement = SolutionModel.incomeStatement;
            incomeStatement.Reverse();

            cashFlow = SolutionModel.cashFlow;
            cashFlow.Reverse();

            for (int i = 0; i < numberOfColumns+4; i++)
            {
                if (i==0)
                {
                    //Income Statement
                    worksheet.Cells[row, col].Value = "Income Statement";
                    worksheet.Cells[row, col].Style.Font.Bold = true;
                    worksheet.Cells[row, col].Style.Font.Color.SetColor(Color.White);

                    //cria o header
                    worksheet.Cells[row + 1, col].Value = "Description (in '000 " + incomestatement[0].ReportedCurrency + ")";
                    worksheet.Cells[row + 1, col].Style.Font.Bold = true;
                    worksheet.Cells[row + 1, col].Style.Font.Color.SetColor(Cores.CorTexto);
                    worksheet.Cells[row + 1, col].Style.Border.Bottom.Style = ExcelBorderStyle.Thin;


                    //Rubricas a analisar da income Statement
                    worksheet.Cells[row + 2, col].Value = "Revenue";
                    worksheet.Cells[row + 3, col].Value = "Gross Profit";
                    worksheet.Cells[row + 4, col].Value = "EBITDA";                  
                    worksheet.Cells[row + 5, col].Value = "Core EBIT";
                    worksheet.Cells[row + 6, col].Value = "NOPLAT";
                    worksheet.Cells[row + 7, col].Value = "Net Income (Attributable to shareholders)";
                    

                    //BS

                    worksheet.Cells[row+9, col].Value = "Balance Sheet";
                    worksheet.Cells[row+9, col].Style.Font.Bold = true;
                    worksheet.Cells[row + 9, col].Style.Font.Color.SetColor(Color.White);


                    worksheet.Cells[row + 10, col].Value = "Description (in '000 " + incomestatement[incomestatement.Count() -1].ReportedCurrency + ")";
                    worksheet.Cells[row + 10, col].Style.Font.Bold = true;
                    worksheet.Cells[row + 10, col].Style.Font.Color.SetColor(Cores.CorTexto);
                    worksheet.Cells[row + 10, col].Style.Border.Bottom.Style = ExcelBorderStyle.Thin;

                    //Rubricas a analisar da BS
                    worksheet.Cells[row + 11, col].Value = "Cash and Cash equivalents";
                    worksheet.Cells[row + 12, col].Value = "Intangible Assets";
                    worksheet.Cells[row + 13, col].Value = "PP&E";
                    worksheet.Cells[row + 14, col].Value = "Net Working Capital (operational)";
                    worksheet.Cells[row + 15, col].Value = "Goodwill";
                    worksheet.Cells[row + 16, col].Value = "Invested Capital";
                    worksheet.Cells[row + 17, col].Value = "Net Debt";
                    worksheet.Cells[row + 18, col].Value = "Equity";
                    worksheet.Cells[row + 19, col].Value = "Retained Earnings";


                    //Cash Flow Statement

                    worksheet.Cells[row + 22, col].Value = "Cash Flow Statement";
                    worksheet.Cells[row + 22, col].Style.Font.Bold = true;
                    worksheet.Cells[row + 22, col].Style.Font.Color.SetColor(Color.White);


                    worksheet.Cells[row + 23, col].Value = "Description (in '000 " + incomestatement[incomestatement.Count() - 1].ReportedCurrency + ")";
                    worksheet.Cells[row + 23, col].Style.Font.Bold = true;
                    worksheet.Cells[row + 23, col].Style.Font.Color.SetColor(Cores.CorTexto);
                    worksheet.Cells[row + 23, col].Style.Border.Bottom.Style = ExcelBorderStyle.Thin;

                    //Rubricas a analisar da BS
                    worksheet.Cells[row + 24, col].Value = "Cash Flow From Operations";
                    worksheet.Cells[row + 25, col].Value = "Cash flow from Financing";
                    worksheet.Cells[row + 26, col].Value = "Cash flow from Investing";
                    worksheet.Cells[row + 27, col].Value = "Operational Free cash flow";
                    worksheet.Columns[col].Width = 40;


                }
                if (i!=numberOfColumns+1)
                {
                    worksheet.Cells[row, col + i].Style.Fill.PatternType = ExcelFillStyle.Solid;
                    worksheet.Cells[row, col + i].Style.Fill.BackgroundColor.SetColor(Cores.CorPrincipal);

                    worksheet.Cells[row + 9, col + i].Style.Fill.PatternType = ExcelFillStyle.Solid;
                    worksheet.Cells[row + 9, col + i].Style.Fill.BackgroundColor.SetColor(Cores.CorPrincipal);

                    worksheet.Cells[row + 22, col + i].Style.Fill.PatternType = ExcelFillStyle.Solid;
                    worksheet.Cells[row + 22, col + i].Style.Fill.BackgroundColor.SetColor(Cores.CorPrincipal);
                }



                if (i>0 && i<numberOfColumns+1)
                {


                    FCFProjections fCFProjections = new FCFProjections();
                    
                    try
                    {
                        years.Add(incomestatement[incomestatement.Count() - numberOfColumns - 1 + i].Date.Year);
                        //Sub-header
                        int aux = numberOfYears - (numberOfColumns-1) + 1 + i;
                        int aux1 = i + 2;
                        string column = nextCol.GetExcelColumnName(aux);
                        string column2 = nextCol.GetExcelColumnName(aux1);
                        worksheet.Cells[row + 1, col + i].Formula = "=YEAR('P&L'!" + column + "3)";
                        worksheet.Cells[row + 1, col + i].Style.HorizontalAlignment = ExcelHorizontalAlignment.Right;
                        worksheet.Cells[row + 1, col + i].Style.Font.Color.SetColor(Cores.CorTexto);
                        worksheet.Cells[row + 1, col + i].Style.Border.Bottom.Style = ExcelBorderStyle.Thin;
                        worksheet.Cells[row + 1, col + i].Style.Font.Bold = true;
                        

                        worksheet.Cells[row + 10, col + i].Formula = "=" + column2 + "3";
                        worksheet.Cells[row + 10, col + i].Style.HorizontalAlignment = ExcelHorizontalAlignment.Right;
                        worksheet.Cells[row + 10, col + i].Style.Font.Color.SetColor(Cores.CorTexto);
                        worksheet.Cells[row + 10, col + i].Style.Border.Bottom.Style = ExcelBorderStyle.Thin;
                        worksheet.Cells[row + 10, col + i].Style.Font.Bold = true;

                        worksheet.Cells[row + 23, col + i].Formula = "=" + column2 + "3";
                        worksheet.Cells[row + 23, col + i].Style.HorizontalAlignment = ExcelHorizontalAlignment.Right;
                        worksheet.Cells[row + 23, col + i].Style.Border.Bottom.Style = ExcelBorderStyle.Thin;
                        worksheet.Cells[row + 23, col + i].Style.Font.Color.SetColor(Cores.CorTexto);
                        worksheet.Cells[row + 23, col + i].Style.Font.Bold = true;

                        
                        worksheet.Cells[row + 2, col + i].Formula = "='P&L'!" + column + "5";
                        worksheet.Cells[row + 2, col + i].Style.Numberformat.Format = "#,##0;(#,##0);-";
                        Revenue.Add((double)incomestatement[numberOfColumns - i].Revenue);
                        worksheet.Cells[row + 3, col + i].Formula = "='P&L'!" + column + "7";
                        worksheet.Cells[row + 3, col + i].Style.Numberformat.Format = "#,##0;(#,##0);-";
                        GrossProfit.Add((double)incomestatement[numberOfColumns - i].GrossProfit);
                        GrossMargin.Add((double)incomestatement[numberOfColumns - i].GrossProfitRatio);
                        worksheet.Cells[row + 4, col + i].Formula = "='P&L'!" + column + "30";
                        worksheet.Cells[row + 4, col + i].Style.Numberformat.Format = "#,##0;(#,##0);-";
                        OperatingMargin.Add((double)incomestatement[numberOfColumns - i].OperatingIncomeRatio);
                        OperatingProfit.Add((double)incomestatement[numberOfColumns - i].OperatingIncome);
                        worksheet.Cells[row + 5, col + i].Formula = "='Valuation Support'!" + column + "14";
                        worksheet.Cells[row + 5, col + i].Style.Numberformat.Format = "#,##0;(#,##0);-";

                        worksheet.Cells[row + 6, col + i].Formula = "='Valuation Support'!" + column + "20";
                        worksheet.Cells[row + 6, col + i].Style.Numberformat.Format = "#,##0;(#,##0);-";
                        double taxRate = (double)(incomeStatement[numberOfColumns - i].Revenue - incomeStatement[numberOfColumns - i].CostOfRevenue - incomeStatement[numberOfColumns - i].OperatingExpenses) * (fCFProjections.GetTaxByYear(tax, LastYear, numberOfColumns, i) / 100);
                        double noplat = (double)incomeStatement[numberOfColumns - i].Revenue - (double)incomeStatement[numberOfColumns - i].CostOfRevenue - (double)incomeStatement[numberOfColumns - i].OperatingExpenses - taxRate;
                        
                        Noplat.Add(noplat);

                        worksheet.Cells[row + 7, col + i].Formula = "='P&L'!" + column + "26";
                        worksheet.Cells[row + 7, col + i].Style.Numberformat.Format = "#,##0;(#,##0);-";


                        worksheet.Cells[row + 11, col + i].Formula = "='BS'!" + column + "7";
                        worksheet.Cells[row + 11, col + i].Style.Numberformat.Format = "#,##0;(#,##0);-";
                        Cash.Add((double)balances[numberOfColumns - i].CashAndCashEquivalents);
                        worksheet.Cells[row + 12, col + i].Formula = "='BS'!" + column + "17";
                        worksheet.Cells[row + 12, col + i].Style.Numberformat.Format = "#,##0;(#,##0);-";
                        worksheet.Cells[row + 13, col + i].Formula = "='BS'!" + column + "15";
                        worksheet.Cells[row + 13, col + i].Style.Numberformat.Format = "#,##0;(#,##0);-";
                        PPE.Add((double)balances[numberOfColumns - i].propertyPlantEquipmentNet);
                        worksheet.Cells[row + 14, col + i].Formula = "='Valuation Support'!" + column + "42";
                        worksheet.Cells[row + 14, col + i].Style.Numberformat.Format = "#,##0;(#,##0);-";

                        if (OpWCAux ==0)
                        {
                            WorkingCapital.Add((double)balanceSheet[numberOfColumns - i].NetReceivables + (double)balanceSheet[numberOfColumns - i].Inventory + (double)balanceSheet[numberOfColumns - i].OtherCurrentAssets - ((double)balanceSheet[numberOfColumns - i].AccountPayables + (double)balanceSheet[numberOfColumns - i].DeferredRevenue + balanceSheet[numberOfColumns - i].OtherCurrentLiabilities));
                        }
                        else
                        {
                            WorkingCapital.Add((double)balanceSheet[numberOfColumns - i].NetReceivables + (double)balanceSheet[numberOfColumns - i].Inventory - ((double)balanceSheet[numberOfColumns - i].AccountPayables + (double)balanceSheet[numberOfColumns - i].DeferredRevenue));
                        }
                        

                        worksheet.Cells[row + 15, col + i].Formula = "='BS'!" + column + "16";
                        worksheet.Cells[row + 15, col + i].Style.Numberformat.Format = "#,##0;(#,##0);-";
                        Goodwill.Add((double)balances[numberOfColumns - i].Goodwill);
                        worksheet.Cells[row + 16, col + i].Formula = "='Valuation Support'!" + column + "45";
                        worksheet.Cells[row + 16, col + i].Style.Numberformat.Format = "#,##0;(#,##0);-";
                        double nwcT = (double)balanceSheet[numberOfColumns - i].NetReceivables + (double)balanceSheet[numberOfColumns - i].Inventory + (double)balanceSheet[numberOfColumns - i].OtherCurrentAssets - (double)(balanceSheet[numberOfColumns - i].AccountPayables + balanceSheet[numberOfColumns - i].DeferredRevenue + balanceSheet[numberOfColumns - i].OtherCurrentLiabilities);
                        double investedCapital = (double)balanceSheet[numberOfColumns - i].propertyPlantEquipmentNet + (double)balanceSheet[numberOfColumns - i].IntangibleAssets + (double)balanceSheet[numberOfColumns - i].Goodwill + nwcT;
                        InvestedCapital.Add(investedCapital);
                        worksheet.Cells[row + 17, col + i].Formula = "='BS'!" + column + "60";
                        worksheet.Cells[row + 17, col + i].Style.Numberformat.Format = "#,##0;(#,##0);-";
                        totalDebt.Add(balances[numberOfColumns - i].TotalDebt);

                        worksheet.Cells[row + 18, col + i].Formula = "='BS'!" + column + "53";
                        worksheet.Cells[row + 18, col + i].Style.Numberformat.Format = "#,##0;(#,##0);-";
                        worksheet.Cells[row + 19, col + i].Formula = "='BS'!" + column + "47";
                        worksheet.Cells[row + 19, col + i].Style.Numberformat.Format = "#,##0;(#,##0);-";
                        RetainedEarnings.Add(balances[numberOfColumns - i].RetainedEarnings / 1000);

                        worksheet.Cells[row + 24, col + i].Formula = "='CFS'!" + column + "11";
                        worksheet.Cells[row + 24, col + i].Style.Numberformat.Format = "#,##0;(#,##0);-";
                        CashFlowFromOperations.Add(cashFlow[numberOfColumns - i].NetCashProvidedByOperatingActivities);
                        worksheet.Cells[row + 25, col + i].Formula = "='CFS'!" + column + "25";
                        worksheet.Cells[row + 25, col + i].Style.Numberformat.Format = "#,##0;(#,##0);-";
                        CashFlowFromFinancing.Add(cashFlow[numberOfColumns - i].NetCashUsedProvidedByFinancingActivities);
                        worksheet.Cells[row + 26, col + i].Formula = "='CFS'!" + column + "18";
                        worksheet.Cells[row + 26, col + i].Style.Numberformat.Format = "#,##0;(#,##0);-";
                        CashFlowFromInvesting.Add(cashFlow[numberOfColumns - i].netCashUsedForInvestingActivites);
                        worksheet.Cells[row + 27, col + i].Formula = "='Valuation Support'!" + column + "29";
                        worksheet.Cells[row + 27, col + i].Style.Numberformat.Format = "#,##0;(#,##0);-";
                        FreeCashFlow.Add(cashFlow[numberOfColumns - i].FreeCashFlow);

                        worksheet.Columns[col + i].Width = 14;
                    }
                    catch (Exception)
                    {

                        
                    }
                    

                    
                }
                else if (i==numberOfColumns+1)
                {
                    worksheet.Columns[col + i].Width = 2;
                }
                else if (i==numberOfColumns+2)
                {
                    int auxcol = numberOfColumns +2;
                    int sparkAux = numberOfColumns + 6;
                    int cagrAux = numberOfColumns - 1;
                    string Lastcolumn = nextCol.GetExcelColumnName(auxcol);
                    string SparkColumn = nextCol.GetExcelColumnName(sparkAux);
                    //IncomeStatement
                    worksheet.Cells[row + 1, col + i].Value = "CAGR";
                    worksheet.Cells[row + 1, col + i].Style.HorizontalAlignment = ExcelHorizontalAlignment.Right;
                    worksheet.Cells[row + 1, col + i].Style.Font.Color.SetColor(Cores.CorTexto);
                    worksheet.Cells[row + 1, col + i].Style.Border.Bottom.Style = ExcelBorderStyle.Thin;
                    worksheet.Cells[row + 1, col + i].Style.Font.Bold = true;
                    for (int a = 0; a < 6; a++)
                    {
                        int rowFormula = row + 2 + a;
                        string aux = '"' + "n.a." + '"';
                        worksheet.Cells[rowFormula, col +i].Formula = "=IFERROR(IF(AND(C" +rowFormula+"<0," + Lastcolumn + rowFormula+"<0),-((" + Lastcolumn + rowFormula + "/C" + rowFormula + ")^(1/" + cagrAux + ")-1),(" + Lastcolumn + rowFormula + "/C" + rowFormula + ")^(1/" + cagrAux + ")-1)," + aux + ")";
                        worksheet.Cells[rowFormula, col + i].Style.Numberformat.Format = "0.00%";
                        var sparklineLine = worksheet.SparklineGroups.Add(eSparklineType.Line, worksheet.Cells[SparkColumn + rowFormula], worksheet.Cells["C" + rowFormula + ":" + Lastcolumn + rowFormula]);
                        worksheet.Column(col + i).Style.HorizontalAlignment = ExcelHorizontalAlignment.Right;
                    }

                    //BalanceSheet
                    worksheet.Cells[row + 10, col + i].Value = "CAGR";
                    worksheet.Cells[row + 10, col + i].Style.HorizontalAlignment = ExcelHorizontalAlignment.Right;
                    worksheet.Cells[row + 10, col + i].Style.Font.Color.SetColor(Cores.CorTexto);
                    worksheet.Cells[row + 10, col + i].Style.Border.Bottom.Style = ExcelBorderStyle.Thin;
                    worksheet.Cells[row + 10, col + i].Style.Font.Bold = true;
                    for (int a = 0; a < 9; a++)
                    {
                        int rowFormula = row + 11 + a;
                        string aux = '"' + "n.a." + '"';
                        worksheet.Cells[rowFormula, col + i].Formula = "=IFERROR(IF(AND(C" + rowFormula + "<0," + Lastcolumn + rowFormula + "<0),-((" + Lastcolumn + rowFormula + "/C" + rowFormula + ")^(1/" + cagrAux + ")-1),(" + Lastcolumn + rowFormula + "/C" + rowFormula + ")^(1/" + cagrAux + ")-1)," + aux + ")";
                        worksheet.Cells[rowFormula, col + i].Style.Numberformat.Format = "0.00%";
                        var sparklineLine = worksheet.SparklineGroups.Add(eSparklineType.Line, worksheet.Cells[SparkColumn + rowFormula], worksheet.Cells["C" + rowFormula + ":" + Lastcolumn + rowFormula]);
                        worksheet.Column(col + i).Style.HorizontalAlignment = ExcelHorizontalAlignment.Right;
                    }

                    //CFS
                    worksheet.Cells[row + 23, col + i].Value = "CAGR";
                    worksheet.Cells[row + 23, col + i].Style.HorizontalAlignment = ExcelHorizontalAlignment.Right;
                    worksheet.Cells[row + 23, col + i].Style.Font.Color.SetColor(Cores.CorTexto);
                    worksheet.Cells[row + 23, col + i].Style.Border.Bottom.Style = ExcelBorderStyle.Thin;
                    worksheet.Cells[row + 23, col + i].Style.Font.Bold = true;
                    for (int a = 0; a < 4; a++)
                    {
                        int rowFormula = row + 24 + a;
                        string aux = '"' + "n.a." + '"';
                        worksheet.Cells[rowFormula, col + i].Formula = "=IFERROR(IF(AND(C" + rowFormula + "<0," + Lastcolumn + rowFormula + "<0),-((" + Lastcolumn + rowFormula + "/C" + rowFormula + ")^(1/" + cagrAux + ")-1),(" + Lastcolumn + rowFormula + "/C" + rowFormula + ")^(1/" + cagrAux + ")-1)," + aux + ")";
                        worksheet.Cells[rowFormula, col + i].Style.Numberformat.Format = "0.00%";
                        var sparklineLine = worksheet.SparklineGroups.Add(eSparklineType.Line, worksheet.Cells[SparkColumn + rowFormula], worksheet.Cells["C" + rowFormula + ":" + Lastcolumn + rowFormula]);
                        worksheet.Column(col + i).Style.HorizontalAlignment = ExcelHorizontalAlignment.Right;
                    }
                    worksheet.Columns[col + i].Width = 14;
                }
                else if (i==numberOfColumns+3)
                {
                    int auxcol = numberOfColumns + 2;
                    string Lastcolumn = nextCol.GetExcelColumnName(auxcol);

                    //IncomeStatement
                    worksheet.Cells[row + 1, col + i].Value = "Coefficient of variation";
                    worksheet.Cells[row + 1, col + i].Style.HorizontalAlignment = ExcelHorizontalAlignment.Right;
                    worksheet.Cells[row + 1, col + i].Style.Font.Color.SetColor(Cores.CorTexto);
                    worksheet.Cells[row + 1, col + i].Style.Border.Bottom.Style = ExcelBorderStyle.Thin;
                    worksheet.Cells[row + 1, col + i].Style.Font.Bold = true;
                    for (int a = 0; a < 6; a++)
                    {
                        int rowFormula = row + 2 + a;
                        string aux = '"' + "n.a." + '"';
                        worksheet.Cells[rowFormula, col + i].Formula = "=IFERROR(STDEV(C" + rowFormula + ":" + Lastcolumn + rowFormula + ")/AVERAGE(C" + rowFormula + ":" + Lastcolumn + rowFormula + ")," + aux + ")";
                        worksheet.Cells[rowFormula, col + i].Style.Numberformat.Format = "0.00%";
                        worksheet.Column(col + i).Style.HorizontalAlignment = ExcelHorizontalAlignment.Right;
                    }

                    //BalanceSheet
                    worksheet.Cells[row + 10, col + i].Value = "Coefficient of variation";
                    worksheet.Cells[row + 10, col + i].Style.HorizontalAlignment = ExcelHorizontalAlignment.Right;
                    worksheet.Cells[row + 10, col + i].Style.Font.Color.SetColor(Cores.CorTexto);
                    worksheet.Cells[row + 10, col + i].Style.Border.Bottom.Style = ExcelBorderStyle.Thin;
                    worksheet.Cells[row + 10, col + i].Style.Font.Bold = true;
                    for (int a = 0; a < 9; a++)
                    {
                        int rowFormula = row + 11 + a;
                        string aux = '"' + "n.a." + '"';
                        worksheet.Cells[rowFormula, col + i].Formula = "=IFERROR(STDEV(C" + rowFormula + ":" + Lastcolumn + rowFormula + ")/AVERAGE(C" + rowFormula + ":" + Lastcolumn + rowFormula + ")," + aux + ")";
                        worksheet.Cells[rowFormula, col + i].Style.Numberformat.Format = "0.00%";
                        worksheet.Column(col + i).Style.HorizontalAlignment = ExcelHorizontalAlignment.Right;
                    }

                    //CFS
                    worksheet.Cells[row + 23, col + i].Value = "Coefficient of variation";
                    worksheet.Cells[row + 23, col + i].Style.HorizontalAlignment = ExcelHorizontalAlignment.Right;
                    worksheet.Cells[row + 23, col + i].Style.Font.Color.SetColor(Cores.CorTexto);
                    worksheet.Cells[row + 23, col + i].Style.Border.Bottom.Style = ExcelBorderStyle.Thin;
                    worksheet.Cells[row + 23, col + i].Style.Font.Bold = true;
                    for (int a = 0; a < 4; a++)
                    {
                        int rowFormula = row + 24 + a;
                        string aux = '"' + "n.a." + '"';
                        worksheet.Cells[rowFormula, col + i].Formula = "=IFERROR(STDEV(C" + rowFormula + ":" + Lastcolumn + rowFormula + ")/AVERAGE(C" + rowFormula + ":" + Lastcolumn + rowFormula + ")," + aux + ")";
                        worksheet.Cells[rowFormula, col + i].Style.Numberformat.Format = "0.00%";
                        worksheet.Column(col + i).Style.HorizontalAlignment = ExcelHorizontalAlignment.Right;
                    }
                    worksheet.Columns[col + i].Width = 25;



                    //Coluna dos comments

                    int commentAux = numberOfColumns +8;
                    string commentCol = nextCol.GetExcelColumnName(commentAux);
                    //IncomeStatement
                    worksheet.Cells[3, numberOfColumns + 7].Value = "Indicator";
                    worksheet.Cells[3, numberOfColumns + 7].Style.Font.Bold = true;
                    worksheet.Cells[3, numberOfColumns + 7].Style.Border.Bottom.Style = ExcelBorderStyle.Thin;
                    worksheet.Columns[numberOfColumns + 7].Width = 10;

                    worksheet.Cells[3, numberOfColumns + 8].Value = "Comments";
                    worksheet.Cells[3, numberOfColumns + 8].Style.Font.Bold = true;
                    worksheet.Cells[3, numberOfColumns + 8].Style.Border.Bottom.Style = ExcelBorderStyle.Thin;
                    worksheet.Columns[numberOfColumns + 8].Width = 50;

                    worksheet.Cells[commentCol + "4:" + commentCol + "8"].Merge = true;
                    worksheet.Cells[row + 2, numberOfColumns + 8].Style.VerticalAlignment = ExcelVerticalAlignment.Top; ;
                    worksheet.Cells[row + 2, numberOfColumns + 8].Style.WrapText = true;
                    
                    
                    CommentBuilderIncomeStatement(worksheet, row + 2, numberOfColumns + 7, Revenue, GrossProfit, GrossMargin, OperatingMargin, OperatingProfit, Noplat);
                    CommentBuilderBalanceSheet(worksheet, row + 12, numberOfColumns + 7, years, WorkingCapital, totalDebt, Cash, PPE, Goodwill, InvestedCapital);
                    CommentBuilderCFS(worksheet, row + 25, numberOfColumns + 7, years, FreeCashFlow, CashFlowFromOperations, CashFlowFromFinancing, CashFlowFromInvesting);

                    //BalanceSheet
                    worksheet.Cells[12, numberOfColumns + 7].Value = "Indicator";
                    worksheet.Cells[12, numberOfColumns + 7].Style.Font.Bold = true;
                    worksheet.Cells[12, numberOfColumns + 7].Style.Border.Bottom.Style = ExcelBorderStyle.Thin;


                    worksheet.Cells[12, numberOfColumns + 8].Value = "Comments";
                    worksheet.Cells[12, numberOfColumns + 8].Style.Font.Bold = true;
                    worksheet.Cells[12, numberOfColumns + 8].Style.Border.Bottom.Style = ExcelBorderStyle.Thin;
                    worksheet.Columns[numberOfColumns + 8].Width = 50;


                    worksheet.Cells[commentCol + "13:" + commentCol + "21"].Merge = true;
                    worksheet.Cells[row + 11, numberOfColumns + 8].Style.VerticalAlignment = ExcelVerticalAlignment.Top; ;
                    worksheet.Cells[row + 11, numberOfColumns + 8].Style.WrapText = true;

                    //CFS
                    worksheet.Cells[25, numberOfColumns + 7].Value = "Indicator";
                    worksheet.Cells[25, numberOfColumns + 7].Style.Font.Bold = true;
                    worksheet.Cells[25, numberOfColumns + 7].Style.Border.Bottom.Style = ExcelBorderStyle.Thin;

                    worksheet.Cells[25, numberOfColumns + 8].Value = "Comments";
                    worksheet.Cells[25, numberOfColumns + 8].Style.Font.Bold = true;
                    worksheet.Cells[25, numberOfColumns + 8].Style.Border.Bottom.Style = ExcelBorderStyle.Thin;
                    worksheet.Columns[numberOfColumns + 8].Width = 50;

                    worksheet.Cells[commentCol + "26:"+ commentCol + "30"].Merge = true;
                    worksheet.Cells[row + 24, numberOfColumns + 8].Style.VerticalAlignment = ExcelVerticalAlignment.Top; ;
                    worksheet.Cells[row + 24, numberOfColumns + 8].Style.WrapText = true;


                }

                //if (i!=0 && i!= 7)
                //{
                //    worksheet.Columns[col + i].Width = 14;
                //}


            }
            


        }
    }
}
