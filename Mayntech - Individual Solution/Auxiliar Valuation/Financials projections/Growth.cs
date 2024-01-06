using Mayntech___Individual_Solution.Auxiliar.Auxiliar.Other;
using Mayntech___Individual_Solution.Auxiliar_Analysis.Auxiliar.Other;
using Mayntech___Individual_Solution.Pages;
using OfficeOpenXml;
using OfficeOpenXml.Style;
using System.Drawing;
using System.Globalization;

namespace Mayntech___Individual_Solution.Auxiliar_Valuation.Financials_projections
{
    public class Growth
    {
        public void GrowthBuilder(ExcelPackage package, int numberOfYears, string solution, List<ROIC> roics)
        {
            ExcelNextCol columnName = new ExcelNextCol();
            var worksheet = package.Workbook.Worksheets.Add("Growth Analysis");
            worksheet.View.ShowGridLines = false;
            worksheet.View.ZoomScale = 80;
            worksheet.View.FreezePanes(4, 3);
            int row = 2;
            int col = 2;
            int numberOfColumns = 0;

            if (solution=="Analysis")
            {
                numberOfColumns = numberOfYears+1;
            }
            else
            {
                numberOfColumns = numberOfYears + 12;
            }

            for (int i = 0; i < numberOfColumns; i++)
            {

                worksheet.Cells[row, col + i].Style.Fill.PatternType = ExcelFillStyle.Solid;
                worksheet.Cells[row, col + i].Style.Fill.BackgroundColor.SetColor(Cores.CorPrincipal);


                if (i == 0)
                {
                    worksheet.Cells[row, col].Value = "Perpetual Growth Rate";
                    worksheet.Cells[row, col].Style.Font.Bold = true;
                    worksheet.Cells[row, col].Style.Font.Color.SetColor(Color.White);

                    worksheet.Cells[row + 1, col + i].Value = "Description";
                    worksheet.Cells[row + 1, col + i].Style.Font.Color.SetColor(Cores.CorTexto);
                    worksheet.Cells[row + 1, col + i].Style.Border.Bottom.Style = ExcelBorderStyle.Thin;
                    worksheet.Cells[row + 1, col + i].Style.Font.Bold = true;

                    worksheet.Cells[row + 3, col + i].Value = "NOPLAT";
                    worksheet.Cells[row + 4, col + i].Value = "Invested Capital";
                    worksheet.Cells[row + 5, col + i].Value = "ROIC";
                    worksheet.Cells[row + 6, col + i].Value = "Net Investment";
                    worksheet.Cells[row + 7, col + i].Value = "RONIC";
                    worksheet.Cells[row + 8, col + i].Value = "Reinvestment Rate";
                    worksheet.Cells[row + 9, col + i].Value = "FCF";
                    worksheet.Cells[row + 10, col + i].Value = "Growth Rate";

                    worksheet.Columns[col + i].Width = 20;
                }
                else
                {
                    string column = columnName.GetExcelColumnName(col + i);
                    string columnLeft = columnName.GetExcelColumnName(col + i - 1);
                    string columnMinusthree = columnName.GetExcelColumnName(col + i - 3);
                    if (solution == "Analysis")
                    {
                        worksheet.Cells[row + 1, col + i].Formula = "=YEAR('P&L'!" + column + "3)";
                    }
                    else
                    {
                        worksheet.Cells[row + 1, col + i].Formula = "='FCF Projections'!" + column + "3";
                    }

                    worksheet.Cells[row + 1, col + i].Style.Font.Color.SetColor(Cores.CorTexto);
                    worksheet.Cells[row + 1, col + i].Style.Border.Bottom.Style = ExcelBorderStyle.Thin;
                    worksheet.Cells[row + 1, col + i].Style.Font.Bold = true;

                    worksheet.Rows[row + 2].Height = 3;

                    //NOPLAT

                    if (solution=="Analysis")
                    {
                        worksheet.Cells[row + 3, col + i].Formula = "='Valuation Support'!" + column + "20";
                    }
                    else
                    {
                        worksheet.Cells[row + 3, col + i].Formula = "='FCF Projections'!" + column + "20";
                    }                    
                    worksheet.Cells[row + 3, col + i].Style.Numberformat.Format = "#,##0;(#,##0);-";

                    //Invested Capital
                    if (solution=="Analysis")
                    {
                        worksheet.Cells[row + 4, col + i].Formula = "='Valuation Support'!" + column + "45";
                    }
                    else
                    {
                        worksheet.Cells[row + 4, col + i].Formula = "='Auxiliar'!" + column + "31";
                    }                 
                    worksheet.Cells[row + 4, col + i].Style.Numberformat.Format = "#,##0;(#,##0);-";

                    //ROIC
                    worksheet.Cells[row + 5, col + i].Formula = "=" + column + "5/" + column + "6";
                    worksheet.Cells[row + 5, col + i].Style.Numberformat.Format = "0.00%";

                    if (i>1)
                    {
                        //Net Investment
                        if (solution == "Analysis")
                        {
                            worksheet.Cells[row + 6, col + i].Formula = "='Valuation Support'!" + column + "47";
                        }
                        else
                        {
                            worksheet.Cells[row + 6, col + i].Formula = "='Auxiliar'!" + column + "33";
                        }
                        
                        worksheet.Cells[row + 6, col + i].Style.Numberformat.Format = "#,##0;(#,##0);-";

                        //RONIC
                        worksheet.Cells[row + 7, col + i].Formula = "=(" + column + "5-" + columnLeft + "5)/(" + column + "6-" + columnLeft + "6)";
                        worksheet.Cells[row + 7, col + i].Style.Numberformat.Format = "0.00%";

                        //Reinvestment Rate
                        worksheet.Cells[row + 8, col + i].Formula = "=" + column + "8/" + column + "5";
                        worksheet.Cells[row + 8, col + i].Style.Numberformat.Format = "0.00%";
                    }

                    //FCF
                    if (solution == "Analysis")
                    {
                        worksheet.Cells[row + 9, col + i].Formula = "='Valuation Support'!" + column + "29";
                    }
                    else
                    {
                        worksheet.Cells[row + 9, col + i].Formula = "='FCF Projections'!" + column + "29";
                    }
                    
                    worksheet.Cells[row + 9, col + i].Style.Numberformat.Format = "#,##0;(#,##0);-";


                    //Growth Rate
                    worksheet.Cells[row + 10, col + i].Formula = "=" + column + "7*" + column + "10";
                    worksheet.Cells[row + 10, col + i].Style.Numberformat.Format = "0.00%";

                    if (i==numberOfYears +11)
                    {
                        worksheet.Cells[row + 10, col + i+1].Formula = "=AVERAGE(" + columnMinusthree + "12:" + column + "12)";
                        worksheet.Cells[row + 10, col + i+1].Style.Numberformat.Format = "0.00%";
                        worksheet.Cells[row + 10, col + i + 1].Style.Font.Bold = true;
                    }

                    if (i>numberOfYears)
                    {
                        for (int a = 0; a < 9; a++)
                        {
                            worksheet.Cells[row + 2 + a, col + i].Style.Fill.PatternType = ExcelFillStyle.Solid;
                            worksheet.Cells[row + 2 + a, col + i].Style.Fill.BackgroundColor.SetColor(Cores.CorCinzentoClaroBackground);
                        }

                    }

                    worksheet.Columns[col + i].Width = 12;

                }
            }

            if (solution == "Analysis" && roics !=null && roics.Count()==1)
            {
                string Lastcolumn = columnName.GetExcelColumnName(col + numberOfColumns-1);

                row += 13;

                NumberFormatInfo provider = new NumberFormatInfo();
                provider.NumberDecimalSeparator = ".";

                double roicLeases = double.Parse(roics[0].ROICwithleases, provider);
                double roicWithoutLeases = double.Parse(roics[0].ROICwithoutleases, provider);

                worksheet.Cells[row + 1, col].Value = "ROIC Analysis";
                worksheet.Cells[row + 1, col].Style.Font.Color.SetColor(Cores.CorTexto);
                worksheet.Cells[row + 1, col].Style.Border.Bottom.Style = ExcelBorderStyle.Thin;
                worksheet.Cells[row + 1, col].Style.Font.Bold = true;

                worksheet.Cells[row + 1, col+1].Style.Font.Color.SetColor(Cores.CorTexto);
                worksheet.Cells[row + 1, col+1].Style.Border.Bottom.Style = ExcelBorderStyle.Thin;
                worksheet.Cells[row + 1, col+1].Style.Font.Bold = true;


                worksheet.Cells[row + 2, col].Value = "Industry ROIC*";
                worksheet.Cells[row + 2, col+1].Value = (roicLeases/100 + roicWithoutLeases/100) /2;
                worksheet.Cells[row + 2, col + 1].Style.Numberformat.Format = "0.00%";

                worksheet.Cells[row + 3, col].Value = "Company Average ROIC";
                worksheet.Cells[row + 3, col + 1].Formula = "=AVERAGE(C7:" + Lastcolumn + "7)";
                worksheet.Cells[row + 3, col + 1].Style.Numberformat.Format = "0.00%";

                worksheet.Cells[row + 4, col].Value = "*Based on Damodaran data: Industry - " + roics[0].IndustryName + "; Location - " + roics[0].Region;
                worksheet.Cells[row + 4, col].Style.Font.Size = 8;
                worksheet.Cells[row + 4, col].Style.Font.Color.SetColor(Color.Gray);

                worksheet.Columns[col].Width = 23;
            }
        }
    }
}
