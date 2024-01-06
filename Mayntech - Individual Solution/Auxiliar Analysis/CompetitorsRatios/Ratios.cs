using Mayntech___Individual_Solution.Auxiliar.Auxiliar.Other;
using Mayntech___Individual_Solution.Pages;
using Microsoft.AspNetCore.SignalR;
using OfficeOpenXml;
using OfficeOpenXml.Drawing.Chart;
using OfficeOpenXml.Style;
using System.Drawing;
using System.Globalization;

namespace Mayntech___Individual_Solution.Auxiliar_Analysis.CompetitorsRatios
{
    public class Ratios : CommentsRatios
    {
        public async void RatioConst(ExcelWorksheet worksheet, string NameAndFormula, string formulaOne,
            string formula2, string formula3, string formula4, string formula5,
            int col, int row, int numberOfYears, IDictionary<string, List<double>> AllCompanyRatios,
            IDictionary<string, double> AllCompaniesAverage,string companyTick, int numberOfYearsIncomeStatement, int Nature)
        {
            ExcelNextCol nextCol = new ExcelNextCol();
            PeersNameAux nameAux = new PeersNameAux();

            worksheet.Cells[row, col].Value = NameAndFormula;
            worksheet.Cells[row, col].Style.Font.Bold = true;
            worksheet.Cells[row, col].Style.Font.Color.SetColor(Color.White);

            int numberOfColumns = Math.Min(numberOfYears, 5);
            for (int i = 0; i < numberOfColumns+3; i++)
            {
                worksheet.Cells[row, col + i].Style.Fill.PatternType = ExcelFillStyle.Solid;
                worksheet.Cells[row, col + i].Style.Fill.BackgroundColor.SetColor(Cores.CorPrincipal);
            }


            // cria o header
            worksheet.Cells[row + 1, col].Value = "Ticker";
            worksheet.Cells[row + 1, col].Style.Font.Bold = true;
            worksheet.Cells[row + 1, col].Style.Font.Color.SetColor(Cores.CorTexto);
            worksheet.Cells[row + 1, col].Style.Border.Bottom.Style = ExcelBorderStyle.Thin;

            //Coluna dos comments
            worksheet.Cells[row + 1, col + numberOfColumns + 4].Value = "Comments";
            worksheet.Cells[row + 1, col + numberOfColumns + 4].Style.Font.Bold = true;
            worksheet.Cells[row + 1, col + numberOfColumns + 4].Style.Border.Bottom.Style = ExcelBorderStyle.Thin;
            worksheet.Columns[col + numberOfColumns + 4].Width = 50;

            int startRow = row + 2;
            int endRow = row + AllCompaniesAverage.Keys.Count()+2;
            worksheet.Cells["L" + startRow + ":L"  + endRow].Merge = true;
            worksheet.Cells[row + 2, col + numberOfColumns + 4].Style.VerticalAlignment = ExcelVerticalAlignment.Top; ;
            worksheet.Cells[row + 2, col + numberOfColumns + 4].Style.WrapText = true;

            //Ordena o dicionário
            var sortedAverageDict = from entry in AllCompaniesAverage orderby entry.Value descending select entry;


            for (int i = 0; i < numberOfColumns; i++)
            {
                int row1 = row;
                int row2 = row;
                worksheet.Cells[row + 1, col + 1 + i].Style.Border.Bottom.Style = ExcelBorderStyle.Thin;
                worksheet.Columns[col + 1 + i].Width = 12;



                int aux = numberOfYears - (numberOfColumns-i) + 3;
                string column = nextCol.GetExcelColumnName(aux);
                worksheet.Cells[row + 1, col + 1 + i].Formula = "=YEAR('P&L'!" + column + "3)";
                worksheet.Cells[row + 1, col + 1 + i].Style.HorizontalAlignment = ExcelHorizontalAlignment.Right;
                worksheet.Cells[row + 1, col + 1 + i].Style.Font.Color.SetColor(Cores.CorTexto);
                worksheet.Cells[row + 1, col + 1 + i].Style.Font.Bold = true;

                foreach (KeyValuePair<string, double> item in sortedAverageDict)
                {
                    try
                    {


                    if (i==0 && item.Key != companyTick)
                    {
                        //string auxName = nameAux.GetCompanyName(item.Key);

                        worksheet.Cells[row2 + 2, col].Value = item.Key;
                        row2++;
                        worksheet.Columns[col].Width = 20;
                    }
                    
                    if (item.Key == companyTick)
                    {
                        worksheet.Cells[row1 + 2, col+i+1].Style.Fill.PatternType = ExcelFillStyle.Solid;
                        worksheet.Cells[row1 + 2, col+i+1].Style.Fill.BackgroundColor.SetColor(Cores.CorBackgroundAnalysis);
                        if (i == 0)
                        {
                            worksheet.Cells[row2 + 2, col].Formula = "='Company Overview'!C4";
                            worksheet.Cells[row2 + 2, col].Style.Fill.PatternType = ExcelFillStyle.Solid;
                            worksheet.Cells[row2 + 2, col].Style.Fill.BackgroundColor.SetColor(Cores.CorBackgroundAnalysis);

                            worksheet.Cells[row2 + 2, col + numberOfColumns + 1].Style.Fill.PatternType = ExcelFillStyle.Solid;
                            worksheet.Cells[row2 + 2, col + numberOfColumns + 1].Style.Fill.BackgroundColor.SetColor(Cores.CorBackgroundAnalysis);

                            worksheet.Cells[row2 + 2, col + numberOfColumns + 2].Style.Fill.PatternType = ExcelFillStyle.Solid;
                            worksheet.Cells[row2 + 2, col + numberOfColumns + 2].Style.Fill.BackgroundColor.SetColor(Cores.CorBackgroundAnalysis);

                            row2++;
                        }
                        
                        if (formula4 == null)
                        {

                                worksheet.Cells[row1 + 2, col + 1 + i].Formula = formulaOne + column + formula2 + column + formula3;
                                if (Nature == 0)
                                {
                                    worksheet.Cells[row1 + 2, col + 1 + i].Style.Numberformat.Format = "0.00%";
                                }
                                else
                                {
                                    worksheet.Cells[row1 + 2, col + 1 + i].Style.Numberformat.Format = "#,##0;(#,##0);-";
                                }

                                row1++;
                        }
                        else if (formula5 == null)
                        {
                            worksheet.Cells[row1 + 2, col + 1 + i].Formula = formulaOne + column + formula2 + column + formula3 + column + formula4;
                                if (Nature == 0)
                                {
                                    worksheet.Cells[row1 + 2, col + 1 + i].Style.Numberformat.Format = "0.00%";
                                }
                                else
                                {
                                    worksheet.Cells[row1 + 2, col + 1 + i].Style.Numberformat.Format = "#,##0;(#,##0);-";
                                }
                                
                            row1++;
                        }
                        else
                        {
                            worksheet.Cells[row1 + 2, col + 1 + i].Formula = formulaOne + column + formula2 + column + formula3 + column + formula4 + column + formula5;
                                if (Nature == 0)
                                {
                                    worksheet.Cells[row1 + 2, col + 1 + i].Style.Numberformat.Format = "0.00%";
                                }
                                else
                                {
                                    worksheet.Cells[row1 + 2, col + 1 + i].Style.Numberformat.Format = "#,##0;(#,##0);-";
                                }
                                
                            row1++;
                        }
                    }
                    else
                    {
                        double valueAux = AllCompanyRatios[item.Key][(numberOfColumns-1) - i];

                        worksheet.Cells[row1 + 2, col + 1 + i].Value = valueAux;
                            if (Nature == 0)
                            {
                                worksheet.Cells[row1 + 2, col + 1 + i].Style.Numberformat.Format = "0.00%";
                            }
                            else
                            {
                                worksheet.Cells[row1 + 2, col + 1 + i].Style.Numberformat.Format = "#,##0;(#,##0);-";

                            }
                        
                        row1++;
                    }
                    }
                    catch 
                    {

                        continue;
                    }
                }

            }
            worksheet.Cells[row + 1, col + numberOfColumns + 1].Value = "Average";
            worksheet.Cells[row + 1, col + numberOfColumns + 2].Value = "Coefficient of Variation";
            worksheet.Cells[row + 1, col + numberOfColumns + 1].Style.Font.Color.SetColor(Cores.CorTexto);
            worksheet.Cells[row + 1, col + numberOfColumns + 1].Style.Font.Bold = true;

            worksheet.Cells[row + 1, col + numberOfColumns + 2].Style.Font.Color.SetColor(Cores.CorTexto);
            worksheet.Cells[row + 1, col + numberOfColumns + 2].Style.Font.Bold = true;

            worksheet.Cells[row + 1, col + numberOfColumns + 1].Style.Border.Bottom.Style = ExcelBorderStyle.Thin;
            worksheet.Cells[row + 1, col + numberOfColumns + 2].Style.Border.Bottom.Style = ExcelBorderStyle.Thin;
            worksheet.Columns[col + numberOfColumns + +2].Width = 25;
            worksheet.Cells[row + 1, col + numberOfColumns + 2].Style.HorizontalAlignment = ExcelHorizontalAlignment.Right;
            worksheet.Columns[col + numberOfColumns + 1].Width = 12;
            worksheet.Cells[row + 1, col + numberOfColumns + 1].Style.HorizontalAlignment = ExcelHorizontalAlignment.Right;

            for (int i = 0; i < AllCompaniesAverage.Keys.Count(); i++)
            {
                int rowFormula = row + 2;
                string column = nextCol.GetExcelColumnName(col + numberOfColumns);
                string error = '"' + "n.a." + '"';
                worksheet.Cells[row + 2, col + numberOfColumns + 1].Formula = "=IFERROR(AVERAGE(C" + rowFormula + ":" + column + rowFormula + ")," + error + ")";
                worksheet.Cells[row + 2, col + numberOfColumns + 2].Formula = "=IFERROR(STDEV(C" + rowFormula + ":"+ column + rowFormula + ")/AVERAGE(C" + rowFormula + ":"+ column + rowFormula + ")," + error + ")";
                if (Nature == 0)
                {
                    worksheet.Cells[row + 2, col + numberOfColumns + 1].Style.Numberformat.Format = "0.00%";
                }
                else
                {
                    worksheet.Cells[row + 2, col + numberOfColumns + 1].Style.Numberformat.Format = "#,##0;(#,##0);-";
                }

                worksheet.Cells[row + 1, col + numberOfColumns + 2].Style.HorizontalAlignment = ExcelHorizontalAlignment.Right;
                worksheet.Cells[row + 2, col + numberOfColumns + 2].Style.HorizontalAlignment = ExcelHorizontalAlignment.Right;
                worksheet.Cells[row + 2, col + numberOfColumns + 2].Style.Numberformat.Format = "0.00%";
                row++;
            }
        }

    }
}
