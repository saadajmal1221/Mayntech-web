using Mayntech___Individual_Solution.Auxiliar.Auxiliar.FMP;
using Mayntech___Individual_Solution.Auxiliar.Auxiliar.Other;
using Mayntech___Individual_Solution.Pages;
using OfficeOpenXml;
using OfficeOpenXml.Sparkline;
using OfficeOpenXml.Style;
using System.Drawing;
using System.Reflection.Metadata.Ecma335;

namespace Mayntech___Individual_Solution.Auxiliar.CompanyAnalysis.WorkingCapital
{
    public class WorkingCapital
    {
        public void WorkingCapitalConstruction(ExcelPackage package, int numberOfYears, string companyName)
        {
            ExcelNextCol nextCol = new ExcelNextCol();
            var worksheet = package.Workbook.Worksheets.Add("Working Capital");
            worksheet.View.ShowGridLines = false;
            worksheet.View.ZoomScale = 80;
            worksheet.View.FreezePanes(4, 3);
            int row = 2;
            int col = 2;


            List<double> cash = new List<double>();
            List<double> shortTermInvestments = new List<double>();
            List<double> AccountsReceivables = new List<double>();
            List<double> Inventory = new List<double>();
            List<double> OtherCurrentAssetLiabilities = new List<double>();
            List<double> AccountsPayable = new List<double>();
            List<double> ShortTermDebt = new List<double>();
            List<double> TaxPayables = new List<double>();
            List<double> DeferredRevenue = new List<double>();
            List<double> WorkingCapital = new List<double>();


            for (int i = 0; i < 9; i++)
            {
                if (i != 6)
                {
                    worksheet.Cells[row, col + i].Style.Fill.PatternType = ExcelFillStyle.Solid;
                    worksheet.Cells[row, col + i].Style.Fill.BackgroundColor.SetColor(Cores.CorPrincipal);
                }


                if (i == 0)
                {
                    worksheet.Cells[row, col].Value = "Working Capital";
                    worksheet.Cells[row, col].Style.Font.Bold = true;
                    worksheet.Cells[row, col].Style.Font.Color.SetColor(Color.White);

                    worksheet.Cells[row + 1, col].Value = "Description (in '000 " + SolutionModel.incomeStatement[0].ReportedCurrency + ")";
                    worksheet.Cells[row + 1, col].Style.Font.Bold = true;
                    worksheet.Cells[row + 1, col].Style.Font.Color.SetColor(Cores.CorTexto);
                    worksheet.Cells[row + 1, col].Style.Border.Bottom.Style = ExcelBorderStyle.Thin;

                    worksheet.Row(row + 2).Height = 4;
                    

                    worksheet.Cells[row + 3, col].Value = companyName;
                    worksheet.Cells[row + 3, col].Style.Font.Bold = true;
                    worksheet.Cells[row + 3, col].Style.Font.Color.SetColor(Cores.CorTexto2);

                    worksheet.Cells[row + 4, col].Value = "Working Capital";
                    worksheet.Cells[row + 4, col].Style.Font.Bold = true;
                    worksheet.Cells[row + 4, col].Style.Font.Color.SetColor(Cores.CorTexto);
                    worksheet.Cells[row + 4, col].Style.Indent = 2;

                    

                    //Rubricas a analisar do WC
                    worksheet.Cells[row + 5, col].Value = "Cash and Cash Equivalents";
                    worksheet.Cells[row + 5, col].Style.Indent = 4;
                    worksheet.Row(row + 5).OutlineLevel = 1;
                    worksheet.Row(row + 5).Collapsed = true;
                    worksheet.Cells[row + 6, col].Value = "Short term investments";
                    worksheet.Cells[row + 6, col].Style.Indent = 4;
                    worksheet.Row(row + 6).OutlineLevel = 1;
                    worksheet.Row(row + 6).Collapsed = true;
                    worksheet.Cells[row + 7, col].Value = "Accounts Receivables";
                    worksheet.Cells[row + 7, col].Style.Indent = 4;
                    worksheet.Row(row + 7).OutlineLevel = 1;
                    worksheet.Row(row + 7).Collapsed = true;
                    worksheet.Cells[row + 8, col].Value = "Inventory";
                    worksheet.Cells[row + 8, col].Style.Indent = 4;
                    worksheet.Row(row + 8).OutlineLevel = 1;
                    worksheet.Row(row + 8).Collapsed = true;
                    worksheet.Cells[row + 9, col].Value = "Other current assets/Liabilities";
                    worksheet.Cells[row + 9, col].Style.Indent = 4;
                    worksheet.Row(row + 9).OutlineLevel = 1;
                    worksheet.Row(row + 9).Collapsed = true;
                    worksheet.Cells[row + 10, col].Value = "Accounts payable";
                    worksheet.Cells[row + 10, col].Style.Indent =4;
                    worksheet.Row(row + 10).OutlineLevel = 1;
                    worksheet.Row(row + 10).Collapsed = true;
                    worksheet.Cells[row + 11, col].Value = "Short term debt";
                    worksheet.Cells[row + 11, col].Style.Indent = 4;
                    worksheet.Row(row + 11).OutlineLevel = 1;
                    worksheet.Row(row + 11).Collapsed = true;
                    worksheet.Cells[row + 12, col].Value = "Tax Payables";
                    worksheet.Cells[row + 12, col].Style.Indent = 4;
                    worksheet.Row(row + 12).OutlineLevel = 1;
                    worksheet.Row(row + 12).Collapsed = true;
                    worksheet.Cells[row + 13, col].Value = "Deferred Revenue";
                    worksheet.Cells[row + 13, col].Style.Indent = 4;
                    worksheet.Row(row + 13).OutlineLevel = 1;
                    worksheet.Row(row + 13).Collapsed = true;

                    worksheet.Columns[col + i].Width = 35;


                    worksheet.Cells[row + 15, col].Value = "Competitors";
                    worksheet.Cells[row + 15, col].Style.Font.Bold = true;
                    worksheet.Cells[row + 15, col].Style.Font.Color.SetColor(Cores.CorTexto2);


                }

                if (i > 0 && i < 6)
                {
                    worksheet.Columns[col + i].Width = 14;

                    //Sub-header
                    int aux = numberOfYears - 4 + 1 + i;
                    int aux1 = i + 2;
                    string column = nextCol.GetExcelColumnName(aux);
                    string column2 = nextCol.GetExcelColumnName(aux1);
                    worksheet.Cells[row + 1, col + i].Formula = "=YEAR('P&L'!" + column + "3)";
                    worksheet.Cells[row + 1, col + i].Style.HorizontalAlignment = ExcelHorizontalAlignment.Right;
                    worksheet.Cells[row + 1, col + i].Style.Font.Color.SetColor(Cores.CorTexto);
                    worksheet.Cells[row + 1, col + i].Style.Border.Bottom.Style = ExcelBorderStyle.Thin;
                    worksheet.Cells[row + 1, col + i].Style.Font.Bold = true;

                    string columnAux = nextCol.GetExcelColumnName(numberOfYears +2 - (5-i));
                    string columnAux1 = nextCol.GetExcelColumnName(2+i);
                    worksheet.Cells[row + 4, col + i].Formula = "=SUM(" + columnAux1 + "7:" + columnAux1 + "15)";
                    worksheet.Cells[row + 4, col + i].Style.Numberformat.Format = "#,##0;(#,##0);-";
                    worksheet.Cells[row + 4, col + i].Style.Font.Bold = true;
                    worksheet.Cells[row + 5, col + i].Style.Numberformat.Format = "#,##0;(#,##0);-";

                    worksheet.Cells[row + 5, col+i].Formula = "=BS!" + columnAux + "7";
                    worksheet.Cells[row + 5, col + i].Style.Numberformat.Format = "#,##0;(#,##0);-";
                    worksheet.Cells[row + 6, col+i].Formula = "=BS!" + columnAux + "8";
                    worksheet.Cells[row + 6, col + i].Style.Numberformat.Format = "#,##0;(#,##0);-";
                    worksheet.Cells[row + 7, col + i].Formula = "=BS!" + columnAux + "9";
                    worksheet.Cells[row + 7, col + i].Style.Numberformat.Format = "#,##0;(#,##0);-";
                    worksheet.Cells[row + 8, col + i].Formula = "=BS!" + columnAux + "10";
                    worksheet.Cells[row + 8, col + i].Style.Numberformat.Format = "#,##0;(#,##0);-";
                    worksheet.Cells[row + 9, col + i].Formula = "=BS!" + columnAux + "11 - BS!"+columnAux + "31";
                    worksheet.Cells[row + 9, col + i].Style.Numberformat.Format = "#,##0;(#,##0);-";
                    worksheet.Cells[row + 10, col + i].Formula = "=-BS!" + columnAux + "27";
                    worksheet.Cells[row + 10, col + i].Style.Numberformat.Format = "#,##0;(#,##0);-";
                    worksheet.Cells[row + 11, col + i].Formula = "=-BS!" + columnAux + "28";
                    worksheet.Cells[row + 11, col + i].Style.Numberformat.Format = "#,##0;(#,##0);-";
                    worksheet.Cells[row + 12, col + i].Formula = "=-BS!" + columnAux + "29";
                    worksheet.Cells[row + 12, col + i].Style.Numberformat.Format = "#,##0;(#,##0);-";
                    worksheet.Cells[row + 13, col + i].Formula = "=-BS!" + columnAux + "30";
                    worksheet.Cells[row + 13, col + i].Style.Numberformat.Format = "#,##0;(#,##0);-";
                }
                else if (i==6)
                {
                    worksheet.Columns[col + i].Width = 2;
                }
                else if (i == 7)
                {

                    // Header
                    worksheet.Cells[row + 1, col + i].Value = "CAGR";
                    worksheet.Cells[row + 1, col + i].Style.HorizontalAlignment = ExcelHorizontalAlignment.Right;
                    worksheet.Cells[row + 1, col + i].Style.Font.Color.SetColor(Cores.CorTexto);
                    worksheet.Cells[row + 1, col + i].Style.Border.Bottom.Style = ExcelBorderStyle.Thin;
                    worksheet.Cells[row + 1, col + i].Style.Font.Bold = true;

                    for (int a = 0; a < 10; a++)
                    {
                        int rowFormula = row + 4 + a;
                        string aux = '"' + "n.a." + '"';
                        worksheet.Cells[row + 4+a, col +i].Formula = "=IFERROR(IF(AND(C" + rowFormula + "<0,G" + rowFormula + "<0),-((G" + rowFormula + "/C" + rowFormula + ")^(1/4)-1),(G" + rowFormula + "/C" + rowFormula + ")^(1/4)-1)," + aux + ")";
                        worksheet.Cells[rowFormula, col + i].Style.Numberformat.Format = "0.00%";
                        
                        worksheet.Cells[rowFormula, col + i].Style.HorizontalAlignment = ExcelHorizontalAlignment.Right;
                        var sparklineLine = worksheet.SparklineGroups.Add(eSparklineType.Line, worksheet.Cells["K" + rowFormula], worksheet.Cells["C" + rowFormula + ":G" + rowFormula]);
                    }
                }

                else if (i == 8)
                {
                    worksheet.Cells[row + 1, col + i].Value = "Coefficient of variation";
                    worksheet.Cells[row + 1, col + i].Style.HorizontalAlignment = ExcelHorizontalAlignment.Right;
                    worksheet.Cells[row + 1, col + i].Style.Font.Color.SetColor(Cores.CorTexto);
                    worksheet.Cells[row + 1, col + i].Style.Border.Bottom.Style = ExcelBorderStyle.Thin;
                    worksheet.Cells[row + 1, col + i].Style.Font.Bold = true;

                    for (int a = 0; a < 10; a++)
                    {
                        int rowFormula = row + 4 + a;
                        string aux = '"' + "n.a." + '"';
                        worksheet.Cells[rowFormula, col + i].Formula = "=IFERROR(STDEV(C" + rowFormula + ":G" + rowFormula + ")/AVERAGE(C" + rowFormula + ":G" + rowFormula + ")," + aux + ")";
                        worksheet.Cells[rowFormula, col + i].Style.Numberformat.Format = "0.00%";
                        
                        worksheet.Cells[rowFormula, col + i].Style.HorizontalAlignment = ExcelHorizontalAlignment.Right;
                    }
                    worksheet.Column(col + i).Width = 24;
                }
            }


        }
    }
}
