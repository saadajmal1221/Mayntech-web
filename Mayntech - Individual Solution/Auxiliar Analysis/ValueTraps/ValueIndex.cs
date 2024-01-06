using Mayntech___Individual_Solution.Auxiliar.Analysis;
using Mayntech___Individual_Solution.Auxiliar.Auxiliar.Other;
using Mayntech___Individual_Solution.Auxiliar.CompanyAnalysis.Competitors;
using Mayntech___Individual_Solution.Pages;
using OfficeOpenXml;
using OfficeOpenXml.Sparkline;
using OfficeOpenXml.Style;
using System;
using System.ComponentModel;
using System.Drawing;

namespace Mayntech___Individual_Solution.Auxiliar.ValueTraps
{
    public class ValueIndex
    {
        public void ValueTrapsConstruction(ExcelPackage package, string companyName)
        {
            ExcelNextCol columnName = new ExcelNextCol();
            var worksheet = package.Workbook.Worksheets.Add("Value Traps");
            worksheet.View.ShowGridLines = false;
            worksheet.View.ZoomScale = 80;
            int row = 2;
            int col = 2;
            worksheet.Cells[row, col].Value = "Value Traps";
            worksheet.Cells[row, col].Style.Font.Bold = true;
            worksheet.Cells[row, col].Style.Font.Color.SetColor(Color.White);
            worksheet.Cells[row, col].Style.Fill.PatternType = ExcelFillStyle.Solid;
            worksheet.Cells[row, col].Style.Fill.BackgroundColor.SetColor(Cores.CorPrincipal);
            row += 3;

            row += CashFlowToNetIncome(worksheet, row);

            List<double> revenues = new List<double>();
            List<double> receivables = new List<double>();
            List<double> CostOfrevenue = new List<double>();
            List<double> payables = new List<double>();
            List<double> inventory = new List<double>();
            for (int i = 0; i < 5; i++)
            {
                try
                {
                    revenues.Add(SolutionModel.incomeStatement[SolutionModel.incomeStatement.Count() - 5 + i].Revenue);
                    receivables.Add(SolutionModel.balances[SolutionModel.balances.Count() - 5 + i].NetReceivables);
                    payables.Add(SolutionModel.balances[SolutionModel.balances.Count() - 5 + i].AccountPayables);
                    CostOfrevenue.Add(SolutionModel.incomeStatement[SolutionModel.incomeStatement.Count() - 5 + i].CostOfRevenue);
                    inventory.Add(SolutionModel.balances[SolutionModel.balances.Count() - 5 + i].Inventory);
                }
                catch (Exception)
                {


                }

            }

            row += Inconsistencies(worksheet, row, revenues, receivables, "Revenue", "Accounts receivable", "P&L", "BS", 5, 9, "Accounts receivable and revenues show some inconsistency.", "There are significant inconsistencies between accounts receivables and revenues. It is crucial to understand the reason behind this.", "Revenues & Receivables inconsistencies");

            row += Inconsistencies(worksheet, row, CostOfrevenue, payables, "Cost of Revenue", "Accounts payable", "P&L", "BS", 6, 27, "Accounts payable and cost of revenue show some inconsistency.", "There are significant inconsistencies between accounts payables and cost of revenue. It is crucial to understand the reason behind this.", "Cost of revenues & Payables inconsistencies");

            row += Inconsistencies(worksheet, row, CostOfrevenue, inventory, "Cost of Revenue", "Inventory", "P&L", "BS", 6, 10, "Inventory and cost of revenue show some inconsistency.", "There are significant inconsistencies between inventory and cost of revenue. It is crucial to understand the reason behind this.", "Cost of revenues & Inventory inconsistencies");
        }

        public int CashFlowToNetIncome(ExcelWorksheet worksheet, int row)
        {
            List<double> netIncome = new List<double>();
            List<double> CashFromOp = new List<double>();
            List<double> xvalues = new List<double>();
            int col = 2;

            ExcelNextCol columnName = new ExcelNextCol();

            for (int i = 0; i < 5; i++)
            {
                try
                {
                    netIncome.Add(SolutionModel.incomeStatement[SolutionModel.incomeStatement.Count() - 5 + i].NetIncome);
                    CashFromOp.Add(SolutionModel.cashFlow[SolutionModel.cashFlow.Count() - 5 + i].NetCashProvidedByOperatingActivities);
                    xvalues.Add(i);
                }
                catch (Exception)
                {


                }

            }

            LinearRegression linearRegression = new LinearRegression();
            List<double> LrNetIncome = linearRegression.LinearRegressionCalculation(xvalues, netIncome);
            List<double> LrCashFromOp = linearRegression.LinearRegressionCalculation(xvalues, CashFromOp);

            CompetitorsAux auxCalc = new CompetitorsAux();
            string evol = auxCalc.Evolution(CashFromOp, netIncome);



            if (evol == "-" || evol == "--" || evol == "+" || evol == "+-" || evol == "++" || evol=="n.a." )
            {
                worksheet.Cells[row - 1, col].Value = "Cash Flow Inconsistency";
                worksheet.Cells[row - 1, col].Style.Font.Bold = true;
                worksheet.Cells[row - 1, col].Style.Font.Color.SetColor(Color.White);

                for (int i = 0; i < 9; i++)
                {
                    worksheet.Cells[row - 1, col + i].Style.Fill.PatternType = ExcelFillStyle.Solid;
                    worksheet.Cells[row - 1, col + i].Style.Fill.BackgroundColor.SetColor(Cores.CorForte4);
                    if (i != 6)
                    {

                        worksheet.Cells[row + 1, col + i].Style.Fill.PatternType = ExcelFillStyle.Solid;
                        worksheet.Cells[row + 1, col + i].Style.Fill.BackgroundColor.SetColor(Cores.CorNumber3);
                        worksheet.Cells[row + 2, col + i].Style.Fill.PatternType = ExcelFillStyle.Solid;
                        worksheet.Cells[row + 2, col + i].Style.Fill.BackgroundColor.SetColor(Cores.CorNumber3);
                        worksheet.Cells[row + 3, col + i].Style.Fill.PatternType = ExcelFillStyle.Solid;
                        worksheet.Cells[row + 3, col + i].Style.Fill.BackgroundColor.SetColor(Cores.CorNumber3);
                        worksheet.Cells[row + 4, col + i].Style.Fill.PatternType = ExcelFillStyle.Solid;
                        worksheet.Cells[row + 4, col + i].Style.Fill.BackgroundColor.SetColor(Cores.CorNumber3);
                    }


                    if (i == 0)
                    {
                        worksheet.Rows[row + 2].Height = 4;

                        worksheet.Cells[row + 1, col+i].Formula = "='P&L'!B3" ;
                        worksheet.Cells[row + 1, col + i].Style.HorizontalAlignment = ExcelHorizontalAlignment.Right;
                        worksheet.Cells[row + 1, col + i].Style.Font.Color.SetColor(Cores.CorTexto);
                        worksheet.Cells[row + 1, col + i].Style.Border.Bottom.Style = ExcelBorderStyle.Thin;
                        worksheet.Cells[row + 1, col + i].Style.Font.Bold = true;

                        worksheet.Cells[row + 3, col + i].Value = "Net Income";
                        worksheet.Cells[row + 3, col + i].Style.Indent = 2;

                        worksheet.Cells[row + 4, col + i].Value = "Cash From Operations";
                        worksheet.Cells[row + 4, col + i].Style.Indent = 2;
                        worksheet.Columns[col + i].Width = 25;

                        worksheet.Cells[row + 1, col + 11].Value = "Comments";
                        worksheet.Cells[row + 1, col + 11].Style.Font.Bold = true;
                        worksheet.Cells[row + 1, col + 11].Style.Border.Bottom.Style = ExcelBorderStyle.Thin;
                        worksheet.Columns[col + 11].Width = 55;

                        if (LrNetIncome[0]>0 && LrCashFromOp[0]<0 || LrNetIncome[2]>0.5 && LrCashFromOp[2]>0.5)
                        {
                            worksheet.Cells[row + 3, col + 10].Style.Fill.PatternType = ExcelFillStyle.Solid;
                            worksheet.Cells[row + 3, col + 10].Style.Fill.BackgroundColor.SetColor(Cores.OrangeWarning);
                            worksheet.Cells[row + 3, col + 11].Value = "Cash from operations is significantly uncorrelated with net income, it is crucial to understand why this is. ";
                        }
                        else if (evol == "-" )
                        {
                            worksheet.Cells[row + 3, col + 10].Style.Fill.PatternType = ExcelFillStyle.Solid;
                            worksheet.Cells[row + 3, col + 10].Style.Fill.BackgroundColor.SetColor(Cores.YellowWarning);
                            worksheet.Cells[row + 3, col + 11].Value = "There are some inconsistencies between net income and cash from operations. It is important to understand what is the reason behind this.";
                        }
                        else if (evol == "--")
                        {
                            worksheet.Cells[row + 3, col + 10].Style.Fill.PatternType = ExcelFillStyle.Solid;
                            worksheet.Cells[row + 3, col + 10].Style.Fill.BackgroundColor.SetColor(Cores.OrangeWarning);
                            worksheet.Cells[row + 3, col + 11].Value = "Cash from operations is significantly uncorrelated with net income, it is crucial to understand why this is. ";
                        }

                        int startRow = row + 3;
                        int endRow = row + 5;
                        worksheet.Cells["M" + startRow + ":M" + endRow].Merge = true;
                        worksheet.Cells["M" + startRow].Style.HorizontalAlignment = ExcelHorizontalAlignment.Left;
                        worksheet.Cells["M" + startRow].Style.VerticalAlignment = ExcelVerticalAlignment.Top;
                        worksheet.Cells["M" + startRow].Style.WrapText = true;

                        worksheet.Cells[row + 1, col + 10].Value = "Indicator";
                        worksheet.Cells[row + 1, col + 10].Style.Font.Bold = true;
                        worksheet.Cells[row + 1, col + 10].Style.Border.Bottom.Style = ExcelBorderStyle.Thin;
                        worksheet.Columns[col + 10].Width = 10;

                    }
                    else if (i > 0 && i < 6)
                    {
                        int aux = SolutionModel.NumberYears - 4 + 1 + i;
                        string column = columnName.GetExcelColumnName(aux);

                        worksheet.Cells[row + 1, col + i].Formula = "=YEAR('P&L'!" + column + "3)";
                        worksheet.Cells[row + 1, col + i].Style.HorizontalAlignment = ExcelHorizontalAlignment.Right;
                        worksheet.Cells[row + 1, col + i].Style.Font.Color.SetColor(Cores.CorTexto);
                        worksheet.Cells[row + 1, col + i].Style.Border.Bottom.Style = ExcelBorderStyle.Thin;
                        worksheet.Cells[row + 1, col + i].Style.Font.Bold = true;



                        worksheet.Cells[row + 3, col + i].Formula = "='P&L'!" + column + "23";
                        worksheet.Cells[row + 3, col + i].Style.Numberformat.Format = "#,##0;(#,##0);-";


                        worksheet.Cells[row + 4, col + i].Formula = "='CFS'!" + column + "11";
                        worksheet.Cells[row + 4, col + i].Style.Numberformat.Format = "#,##0;(#,##0);-";

                        worksheet.Columns[col + i].Width = 15;

                    }
                    else if (i == 6)
                    {
                        worksheet.Columns[col + i].Width = 2;
                    }
                    else if (i == 7)
                    {
                        worksheet.Cells[row + 1, col + i].Value = "CAGR";
                        worksheet.Cells[row + 1, col + i].Style.HorizontalAlignment = ExcelHorizontalAlignment.Right;
                        worksheet.Cells[row + 1, col + i].Style.Font.Color.SetColor(Cores.CorTexto);
                        worksheet.Cells[row + 1, col + i].Style.Border.Bottom.Style = ExcelBorderStyle.Thin;
                        worksheet.Cells[row + 1, col + i].Style.Font.Bold = true;

                        for (int a = 0; a < 2; a++)
                        {
                            int rowFormula = row + 3 + a;
                            string aux = '"' + "n.a." + '"';
                            worksheet.Cells[rowFormula, col + i].Formula = "=IFERROR(IF(AND(C" + rowFormula + "<0,G" + rowFormula + "<0),-((G" + rowFormula + "/C" + rowFormula + ")^(1/4)-1),(G" + rowFormula + "/C" + rowFormula + ")^(1/4)-1)," + aux + ")";
                            worksheet.Cells[rowFormula, col + i].Style.Numberformat.Format = "0.00%";
                            var sparklineLine = worksheet.SparklineGroups.Add(eSparklineType.Line, worksheet.Cells["K" + rowFormula], worksheet.Cells["C" + rowFormula + ":G" + rowFormula]);
                            worksheet.Column(col + i).Style.HorizontalAlignment = ExcelHorizontalAlignment.Right;
                        }
                        worksheet.Columns[col + i].Width = 10;
                    }
                    else if (i == 8)
                    {
                        worksheet.Cells[row + 1, col + i].Value = "Coefficient of Variation";
                        worksheet.Cells[row + 1, col + i].Style.HorizontalAlignment = ExcelHorizontalAlignment.Right;
                        worksheet.Cells[row + 1, col + i].Style.Font.Color.SetColor(Cores.CorTexto);
                        worksheet.Cells[row + 1, col + i].Style.Border.Bottom.Style = ExcelBorderStyle.Thin;
                        worksheet.Cells[row + 1, col + i].Style.Font.Bold = true;

                        for (int a = 0; a < 2; a++)
                        {
                            int rowFormula = row + 3 + a;
                            string aux = '"' + "n.a." + '"';
                            worksheet.Cells[rowFormula, col + i].Formula = "=IFERROR(STDEV(C" + rowFormula + ":G" + rowFormula + ")/AVERAGE(C" + rowFormula + ":G" + rowFormula + ")," + aux + ")";
                            worksheet.Cells[rowFormula, col + i].Style.Numberformat.Format = "0.00%";
                            worksheet.Column(col + i).Style.HorizontalAlignment = ExcelHorizontalAlignment.Right;
                        }
                        worksheet.Columns[col + i].Width = 24;
                    }

                }
                return 9;
            }
            return 0;
        }

        public int Inconsistencies(ExcelWorksheet worksheet, int row,  List<double> valuesOne, List<double> ValuesTwo, string NameOfFirstRow,
            string NameOfSecondRow, string SheetOfFirstVariable, string sheetOfSecondVariable, int NumFirstVariable,
            int NumSecondVariable, string CommentYellow, string CommentRed, string NameOfAnalysis)
        {

            int col = 2;

            ExcelNextCol columnName = new ExcelNextCol();
            List<double> xvalues = new List<double>();
            for (int i = 0; i < valuesOne.Count(); i++)
            {
                xvalues.Add(i);
            }


            LinearRegression linearregression = new LinearRegression();
            List<double> output = linearregression.LinearRegressionCalculation(ValuesTwo, valuesOne);
            List<double> FirstNumberLr = linearregression.LinearRegressionCalculation(xvalues, valuesOne);
            List<double> SecondNumberLr = linearregression.LinearRegressionCalculation(xvalues, ValuesTwo);
            string inconsistency = null;

            if (FirstNumberLr[2]>0.4 && SecondNumberLr[2]>0.4)
            {
                if (FirstNumberLr[0] * SecondNumberLr[0] < 0)
                {
                    inconsistency = "yes";
                }
            }

            if (output[2]<0.3 || output[0]<0 || inconsistency == "yes")
            {

                worksheet.Cells[row - 1, col].Value = NameOfAnalysis;
                worksheet.Cells[row - 1, col].Style.Font.Bold = true;
                worksheet.Cells[row - 1, col].Style.Font.Color.SetColor(Color.White);

                for (int i = 0; i < 9; i++)
                {
                    worksheet.Cells[row - 1, col + i].Style.Fill.PatternType = ExcelFillStyle.Solid;
                    worksheet.Cells[row - 1, col + i].Style.Fill.BackgroundColor.SetColor(Cores.CorForte4);
                    if (i != 6)
                    {

                        worksheet.Cells[row + 1, col + i].Style.Fill.PatternType = ExcelFillStyle.Solid;
                        worksheet.Cells[row + 1, col + i].Style.Fill.BackgroundColor.SetColor(Cores.CorNumber3);
                        worksheet.Cells[row + 2, col + i].Style.Fill.PatternType = ExcelFillStyle.Solid;
                        worksheet.Cells[row + 2, col + i].Style.Fill.BackgroundColor.SetColor(Cores.CorNumber3);
                        worksheet.Cells[row + 3, col + i].Style.Fill.PatternType = ExcelFillStyle.Solid;
                        worksheet.Cells[row + 3, col + i].Style.Fill.BackgroundColor.SetColor(Cores.CorNumber3);
                        worksheet.Cells[row + 4, col + i].Style.Fill.PatternType = ExcelFillStyle.Solid;
                        worksheet.Cells[row + 4, col + i].Style.Fill.BackgroundColor.SetColor(Cores.CorNumber3);
                    }


                    if (i == 0)
                    {
                        worksheet.Rows[row + 2].Height = 4;

                        worksheet.Cells[row + 1, col + i].Formula = "='P&L'!B3";
                        worksheet.Cells[row + 1, col + i].Style.HorizontalAlignment = ExcelHorizontalAlignment.Right;
                        worksheet.Cells[row + 1, col + i].Style.Font.Color.SetColor(Cores.CorTexto);
                        worksheet.Cells[row + 1, col + i].Style.Border.Bottom.Style = ExcelBorderStyle.Thin;
                        worksheet.Cells[row + 1, col + i].Style.Font.Bold = true;

                        worksheet.Cells[row + 3, col + i].Value = NameOfFirstRow;
                        worksheet.Cells[row + 3, col + i].Style.Indent = 2;

                        worksheet.Cells[row + 4, col + i].Value = NameOfSecondRow;
                        worksheet.Cells[row + 4, col + i].Style.Indent = 2;
                        worksheet.Columns[col + i].Width = 25;

                        worksheet.Cells[row + 1, col + 11].Value = "Comments";
                        worksheet.Cells[row + 1, col + 11].Style.Font.Bold = true;
                        worksheet.Cells[row + 1, col + 11].Style.Border.Bottom.Style = ExcelBorderStyle.Thin;
                        worksheet.Columns[col + 11].Width = 55;

                        if (output[2] < 0.05 && inconsistency == "yes")
                        {
                            worksheet.Cells[row + 3, col + 10].Style.Fill.PatternType = ExcelFillStyle.Solid;
                            worksheet.Cells[row + 3, col + 10].Style.Fill.BackgroundColor.SetColor(Cores.OrangeWarning);
                            worksheet.Cells[row + 3, col + 11].Value = CommentRed;
                        }
                        else if (output[2] < 0.05)
                        {
                            worksheet.Cells[row + 3, col + 10].Style.Fill.PatternType = ExcelFillStyle.Solid;
                            worksheet.Cells[row + 3, col + 10].Style.Fill.BackgroundColor.SetColor(Cores.YellowWarning);
                            worksheet.Cells[row + 3, col + 11].Value = CommentYellow;
                        }
                        else if (output[0]<0)
                        {
                            worksheet.Cells[row + 3, col + 10].Style.Fill.PatternType = ExcelFillStyle.Solid;
                            worksheet.Cells[row + 3, col + 10].Style.Fill.BackgroundColor.SetColor(Cores.OrangeWarning);
                            worksheet.Cells[row + 3, col + 11].Value = CommentRed;
                        }
                        else
                        {
                            worksheet.Cells[row + 3, col + 10].Style.Fill.PatternType = ExcelFillStyle.Solid;
                            worksheet.Cells[row + 3, col + 10].Style.Fill.BackgroundColor.SetColor(Color.DarkGray);
                            worksheet.Cells[row + 3, col + 11].Value = CommentYellow;
                        }

                        int startRow = row + 3;
                        int endRow = row + 5;
                        worksheet.Cells["M" + startRow + ":M" + endRow].Merge = true;
                        worksheet.Cells["M" + startRow].Style.HorizontalAlignment = ExcelHorizontalAlignment.Left;
                        worksheet.Cells["M" + startRow].Style.VerticalAlignment = ExcelVerticalAlignment.Top;
                        worksheet.Cells["M" + startRow].Style.WrapText = true;

                        worksheet.Cells[row + 1, col + 10].Value = "Indicator";
                        worksheet.Cells[row + 1, col + 10].Style.Font.Bold = true;
                        worksheet.Cells[row + 1, col + 10].Style.Border.Bottom.Style = ExcelBorderStyle.Thin;
                        worksheet.Columns[col + 10].Width = 10;

                    }
                    else if (i > 0 && i < 6)
                    {
                        int aux = SolutionModel.NumberYears - 4 + 1 + i;
                        string column = columnName.GetExcelColumnName(aux);

                        worksheet.Cells[row + 1, col + i].Formula = "=YEAR('P&L'!" + column + "3)";
                        worksheet.Cells[row + 1, col + i].Style.HorizontalAlignment = ExcelHorizontalAlignment.Right;
                        worksheet.Cells[row + 1, col + i].Style.Font.Color.SetColor(Cores.CorTexto);
                        worksheet.Cells[row + 1, col + i].Style.Border.Bottom.Style = ExcelBorderStyle.Thin;
                        worksheet.Cells[row + 1, col + i].Style.Font.Bold = true;


                        if (NameOfFirstRow == "Cost of Revenue")
                        {
                            worksheet.Cells[row + 3, col + i].Formula = "=-'" + SheetOfFirstVariable + "'!" + column + NumFirstVariable;
                            worksheet.Cells[row + 3, col + i].Style.Numberformat.Format = "#,##0;(#,##0);-";
                        }
                        else
                        {
                            worksheet.Cells[row + 3, col + i].Formula = "='" + SheetOfFirstVariable + "'!" + column + NumFirstVariable;
                            worksheet.Cells[row + 3, col + i].Style.Numberformat.Format = "#,##0;(#,##0);-";
                        }



                        worksheet.Cells[row + 4, col + i].Formula = "='" + sheetOfSecondVariable + "'!" + column + NumSecondVariable;
                        worksheet.Cells[row + 4, col + i].Style.Numberformat.Format = "#,##0;(#,##0);-";

                        worksheet.Columns[col + i].Width = 15;

                    }
                    else if (i == 6)
                    {
                        worksheet.Columns[col + i].Width = 2;
                    }
                    else if (i == 7)
                    {
                        worksheet.Cells[row + 1, col + i].Value = "CAGR";
                        worksheet.Cells[row + 1, col + i].Style.HorizontalAlignment = ExcelHorizontalAlignment.Right;
                        worksheet.Cells[row + 1, col + i].Style.Font.Color.SetColor(Cores.CorTexto);
                        worksheet.Cells[row + 1, col + i].Style.Border.Bottom.Style = ExcelBorderStyle.Thin;
                        worksheet.Cells[row + 1, col + i].Style.Font.Bold = true;

                        for (int a = 0; a < 2; a++)
                        {
                            int rowFormula = row + 3 + a;
                            string aux = '"' + "n.a." + '"';
                            worksheet.Cells[rowFormula, col + i].Formula = "=IFERROR(IF(AND(C" + rowFormula + "<0,G" + rowFormula + "<0),-((G" + rowFormula + "/C" + rowFormula + ")^(1/4)-1),(G" + rowFormula + "/C" + rowFormula + ")^(1/4)-1)," + aux + ")";
                            worksheet.Cells[rowFormula, col + i].Style.Numberformat.Format = "0.00%";
                            var sparklineLine = worksheet.SparklineGroups.Add(eSparklineType.Line, worksheet.Cells["K" + rowFormula], worksheet.Cells["C" + rowFormula + ":G" + rowFormula]);
                            worksheet.Column(col + i).Style.HorizontalAlignment = ExcelHorizontalAlignment.Right;
                        }
                        worksheet.Columns[col + i].Width = 10;
                    }
                    else if (i == 8)
                    {
                        worksheet.Cells[row + 1, col + i].Value = "Coefficient of Variation";
                        worksheet.Cells[row + 1, col + i].Style.HorizontalAlignment = ExcelHorizontalAlignment.Right;
                        worksheet.Cells[row + 1, col + i].Style.Font.Color.SetColor(Cores.CorTexto);
                        worksheet.Cells[row + 1, col + i].Style.Border.Bottom.Style = ExcelBorderStyle.Thin;
                        worksheet.Cells[row + 1, col + i].Style.Font.Bold = true;

                        for (int a = 0; a < 2; a++)
                        {
                            int rowFormula = row + 3 + a;
                            string aux = '"' + "n.a." + '"';
                            worksheet.Cells[rowFormula, col + i].Formula = "=IFERROR(STDEV(C" + rowFormula + ":G" + rowFormula + ")/AVERAGE(C" + rowFormula + ":G" + rowFormula + ")," + aux + ")";
                            worksheet.Cells[rowFormula, col + i].Style.Numberformat.Format = "0.00%";
                            worksheet.Column(col + i).Style.HorizontalAlignment = ExcelHorizontalAlignment.Right;
                        }
                        worksheet.Columns[col + i].Width = 24;
                    }

                }

                return 9;
            }
            return 0;
        }
    }
}
