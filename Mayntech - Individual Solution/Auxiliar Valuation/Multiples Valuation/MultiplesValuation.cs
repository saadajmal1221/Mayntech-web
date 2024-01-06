using Mayntech___Individual_Solution.Auxiliar.Auxiliar.FMP;
using Mayntech___Individual_Solution.Auxiliar.Auxiliar.Other;
using OfficeOpenXml;
using OfficeOpenXml.Style;
using System.Drawing;

namespace Mayntech___Individual_Solution.Auxiliar_Valuation.Multiples_Valuation
{
    public class MultiplesValuation
    {
        public void MultiplesValuationConstruction(ExcelPackage package, int year, int numberOfYears, List<PeersProfile> peersOutlook)
        {
            ExcelNextCol columnName = new ExcelNextCol();
            var worksheet = package.Workbook.Worksheets.Add("Multiples Valuation");
            worksheet.View.ShowGridLines = false;
            worksheet.View.ZoomScale = 80;

            int row = 2;
            int col = 2;

            for (int i = 0; i < 2; i++)
            {
                worksheet.Cells[row, col + i].Style.Fill.PatternType = ExcelFillStyle.Solid;
                worksheet.Cells[row, col + i].Style.Fill.BackgroundColor.SetColor(Cores.CorPrincipal);

                if (i == 0)
                {
                    worksheet.Columns[col + i].Width = 25;

                    worksheet.Cells[row, col + i].Value = "Company Info " + year;
                    worksheet.Cells[row, col + i].Style.Font.Bold = true;
                    worksheet.Cells[row, col + i].Style.Font.Color.SetColor(Color.White);

                    worksheet.Cells[row + 1, col + i].Value = "Description";
                    worksheet.Cells[row + 1, col + i].Style.Font.Color.SetColor(Cores.CorTexto);
                    worksheet.Cells[row + 1, col + i].Style.Border.Bottom.Style = ExcelBorderStyle.Thin;
                    worksheet.Cells[row + 1, col + i].Style.Font.Bold = true;

                    worksheet.Rows[row + 2].Height = 3;

                    worksheet.Cells[row + 3, col + i].Value = "Sales";
                    worksheet.Cells[row + 4, col + i].Value = "EBITDA";
                    worksheet.Cells[row + 5, col + i].Value = "EBIT";
                    worksheet.Cells[row + 6, col + i].Value = "NOPLAT";
                    worksheet.Cells[row + 7, col + i].Value = "Earnings";
                    worksheet.Cells[row + 8, col + i].Value = "Equity";


                    worksheet.Cells[row + 12, col + i].Value = "Equity Value";
                    worksheet.Cells[row+12, col + i].Style.Fill.PatternType = ExcelFillStyle.Solid;
                    worksheet.Cells[row+12, col + i].Style.Fill.BackgroundColor.SetColor(Cores.CorBackgroundAnalysisRevenue);

                    worksheet.Cells[row + 13, col + i].Value = "Net Debt";
                    worksheet.Cells[row + 13, col + i].Style.Fill.PatternType = ExcelFillStyle.Solid;
                    worksheet.Cells[row + 13, col + i].Style.Fill.BackgroundColor.SetColor(Cores.CorCinzentoClaroBackground);

                    worksheet.Cells[row + 14, col + i].Value = "Minority Interests";
                    worksheet.Cells[row + 14, col + i].Style.Fill.PatternType = ExcelFillStyle.Solid;
                    worksheet.Cells[row + 14, col + i].Style.Fill.BackgroundColor.SetColor(Cores.CorCinzentoClaroBackground);

                    worksheet.Cells[row + 15, col + i].Value = "Price Per share";
                    worksheet.Cells[row + 15, col + i].Style.Fill.PatternType = ExcelFillStyle.Solid;
                    worksheet.Cells[row + 15, col + i].Style.Fill.BackgroundColor.SetColor(Cores.CorBackgroundAnalysisRevenue);

                }
                else if (i == 1)
                {
                    worksheet.Cells[row + 1, col + i].Value = "('000)";
                    worksheet.Cells[row + 1, col + i].Style.HorizontalAlignment = ExcelHorizontalAlignment.Right;
                    worksheet.Cells[row + 1, col + i].Style.Font.Color.SetColor(Cores.CorTexto);
                    worksheet.Cells[row + 1, col + i].Style.Border.Bottom.Style = ExcelBorderStyle.Thin;
                    worksheet.Cells[row + 1, col + i].Style.Font.Bold = true;

                    string column = columnName.GetExcelColumnName(numberOfYears + 2);
                    worksheet.Cells[row + 3, col + i].Formula = "='P&L'!" + column + "5";
                    worksheet.Cells[row + 4, col + i].Formula = "='P&L'!" + column + "30";
                    worksheet.Cells[row + 5, col + i].Formula = "='P&L'!" + column + "30+'P&L'!" + column + "32 +'P&L'!" + column + "33+'P&L'!" + column + "34";
                    worksheet.Cells[row + 6, col + i].Formula = "='FCF Projections'!" + column + "20";
                    worksheet.Cells[row + 7, col + i].Formula = "='P&L'!" + column + "23";
                    worksheet.Cells[row + 8, col + i].Formula = "='BS'!" + column + "53";

                    for (int a = 0; a < 6; a++)
                    {
                        worksheet.Cells[row + 3 + a, col + i].Style.Numberformat.Format = "#,##0;(#,##0);-";
                    }

                    worksheet.Cells[row + 12, col + i].Formula = "=AVERAGE(P13:P16)";
                    worksheet.Cells[row + 12, col + i].Style.Fill.PatternType = ExcelFillStyle.Solid;
                    worksheet.Cells[row + 12, col + i].Style.Fill.BackgroundColor.SetColor(Cores.CorBackgroundAnalysisRevenue);
                    worksheet.Cells[row + 12, col + i].Style.Numberformat.Format = "#,##0;(#,##0);-";

                    worksheet.Cells[row + 13, col + i].Formula = "='DCF Valuation'!C10/1000";
                    worksheet.Cells[row + 13, col + i].Style.Fill.PatternType = ExcelFillStyle.Solid;
                    worksheet.Cells[row + 13, col + i].Style.Fill.BackgroundColor.SetColor(Cores.CorCinzentoClaroBackground);
                    worksheet.Cells[row + 13, col + i].Style.Numberformat.Format = "#,##0;(#,##0);-";

                    worksheet.Cells[row + 14, col + i].Formula = "='DCF Valuation'!C11/1000";
                    worksheet.Cells[row + 14, col + i].Style.Fill.PatternType = ExcelFillStyle.Solid;
                    worksheet.Cells[row + 14, col + i].Style.Fill.BackgroundColor.SetColor(Cores.CorCinzentoClaroBackground);
                    worksheet.Cells[row + 14, col + i].Style.Numberformat.Format = "#,##0;(#,##0);-";

                    worksheet.Cells[row + 15, col + i].Formula = "=(C14*1000 - C15*1000 - C16*1000)/'DCF Valuation'!C14";
                    worksheet.Cells[row + 15, col + i].Style.Fill.PatternType = ExcelFillStyle.Solid;
                    worksheet.Cells[row + 15, col + i].Style.Fill.BackgroundColor.SetColor(Cores.CorBackgroundAnalysisRevenue);
                    worksheet.Cells[row + 15, col + i].Style.Numberformat.Format = "#,##0;(#,##0);-";

                    worksheet.Columns[col+i].Width = 12;
                }

            }
            col += 4;

            for (int i = 0; i < 5; i++)
            {
                worksheet.Cells[row, col + i].Style.Fill.PatternType = ExcelFillStyle.Solid;
                worksheet.Cells[row, col + i].Style.Fill.BackgroundColor.SetColor(Cores.CorPrincipal);

                if (i == 0)
                {
                    worksheet.Cells[row, col + i].Value = "Multiples";
                    worksheet.Cells[row, col + i].Style.Font.Bold = true;
                    worksheet.Cells[row, col + i].Style.Font.Color.SetColor(Color.White);

                    worksheet.Cells[row + 1, col].Value = "Company Name";
                    worksheet.Cells[row + 1, col + i].Style.Font.Color.SetColor(Cores.CorTexto);
                    worksheet.Cells[row + 1, col + i].Style.Border.Bottom.Style = ExcelBorderStyle.Thin;
                    worksheet.Cells[row + 1, col + i].Style.Font.Bold = true;
                }
                else if (i == 1)
                {
                    worksheet.Cells[row + 1, col + i].Value = "EV/Book";
                    worksheet.Cells[row + 1, col + i].Style.Font.Color.SetColor(Cores.CorTexto);
                    worksheet.Cells[row + 1, col + i].Style.Border.Bottom.Style = ExcelBorderStyle.Thin;
                    worksheet.Cells[row + 1, col + i].Style.Font.Bold = true;
                    worksheet.Cells[row + 1, col + i].Style.HorizontalAlignment = ExcelHorizontalAlignment.Right;
                }
                else if (i == 2)
                {
                    worksheet.Cells[row + 1, col + i].Value = "EV/EBITDA";
                    worksheet.Cells[row + 1, col + i].Style.Font.Color.SetColor(Cores.CorTexto);
                    worksheet.Cells[row + 1, col + i].Style.Border.Bottom.Style = ExcelBorderStyle.Thin;
                    worksheet.Cells[row + 1, col + i].Style.Font.Bold = true;
                    worksheet.Cells[row + 1, col + i].Style.HorizontalAlignment = ExcelHorizontalAlignment.Right;
                }
                else if (i == 3)
                {
                    worksheet.Cells[row + 1, col + i].Value = "EV/EBIT";
                    worksheet.Cells[row + 1, col + i].Style.Font.Color.SetColor(Cores.CorTexto);
                    worksheet.Cells[row + 1, col + i].Style.Border.Bottom.Style = ExcelBorderStyle.Thin;
                    worksheet.Cells[row + 1, col + i].Style.Font.Bold = true;
                    worksheet.Cells[row + 1, col + i].Style.HorizontalAlignment = ExcelHorizontalAlignment.Right;
                }
                else if (i == 4)
                {
                    worksheet.Cells[row + 1, col + i].Value = "EV/SALES";
                    worksheet.Cells[row + 1, col + i].Style.Font.Color.SetColor(Cores.CorTexto);
                    worksheet.Cells[row + 1, col + i].Style.Border.Bottom.Style = ExcelBorderStyle.Thin;
                    worksheet.Cells[row + 1, col + i].Style.Font.Bold = true;
                    worksheet.Cells[row + 1, col + i].Style.HorizontalAlignment = ExcelHorizontalAlignment.Right;
                }

            }
            worksheet.Columns[col].Width = 25;
            worksheet.Columns[col + 1].Width = 12;

            for (int peers = 0; peers < peersOutlook.Count(); peers++)
            {
                for (int i = 0; i < 5; i++)
                {
                    try
                    {
                        if (i == 0)
                        {
                            worksheet.Cells[row + 3 + peers, col + i].Value = peersOutlook[peers].profile.symbol;
                        }
                        else if (i == 1)
                        {
                            worksheet.Cells[row + 3 + peers, col + i].Value = (double)peersOutlook[peers].profile.mktCap / peersOutlook[peers].financialsAnnual.balance[0].TotalEquity;
                            worksheet.Cells[row + 3 + peers, col + i].Style.Numberformat.Format = "#,##0.00;(#,##0.00);-";
                            worksheet.Columns[col + i].Width = 12;
                        }
                        else if (i == 2)
                        {
                            worksheet.Cells[row + 3 + peers, col + i].Value = ((double)peersOutlook[peers].profile.mktCap + peersOutlook[peers].financialsAnnual.balance[0].TotalDebt - peersOutlook[peers].financialsAnnual.balance[0].CashAndCashEquivalents) / peersOutlook[peers].financialsAnnual.income[0].EBITDA;
                            worksheet.Cells[row + 3 + peers, col + i].Style.Numberformat.Format = "#,##0.00;(#,##0.00);-";
                            worksheet.Columns[col + i].Width = 12;
                        }
                        else if (i == 3)
                        {
                            worksheet.Cells[row + 3 + peers, col + i].Value = ((double)peersOutlook[peers].profile.mktCap + peersOutlook[peers].financialsAnnual.balance[0].TotalDebt - peersOutlook[peers].financialsAnnual.balance[0].CashAndCashEquivalents) / peersOutlook[peers].financialsAnnual.income[0].EBITDA;
                            worksheet.Cells[row + 3 + peers, col + i].Style.Numberformat.Format = "#,##0.00;(#,##0.00);-";
                            worksheet.Columns[col + i].Width = 12;
                        }
                        else if (i == 4)
                        {
                            worksheet.Cells[row + 3 + peers, col + i].Value = ((double)peersOutlook[peers].profile.mktCap + peersOutlook[peers].financialsAnnual.balance[0].TotalDebt - peersOutlook[peers].financialsAnnual.balance[0].CashAndCashEquivalents) / peersOutlook[peers].financialsAnnual.income[0].Revenue;
                            worksheet.Cells[row + 3 + peers, col + i].Style.Numberformat.Format = "#,##0.00;(#,##0.00);-";
                            worksheet.Columns[col + i].Width = 12;
                        }
                    }
                    catch 
                    {

                        continue;
                    }
                    
                }

            }


            col += 7;
            //Tabela Resumo
            int row1 = row;

            for (int i = 0; i < 6; i++)
            {
                worksheet.Cells[row1, col + i].Style.Fill.PatternType = ExcelFillStyle.Solid;
                worksheet.Cells[row1, col + i].Style.Fill.BackgroundColor.SetColor(Cores.CorPrincipal);

                if (i == 0)
                {
                    worksheet.Cells[row1, col + i].Value = "Summary";
                    worksheet.Cells[row1, col + i].Style.Font.Bold = true;
                    worksheet.Cells[row1, col + i].Style.Font.Color.SetColor(Color.White);

                    worksheet.Cells[row1 + 1, col].Value = "Multiples";
                    worksheet.Cells[row1+ 1, col + i].Style.Font.Color.SetColor(Cores.CorTexto);
                    worksheet.Cells[row1 + 1, col + i].Style.Border.Bottom.Style = ExcelBorderStyle.Thin;
                    worksheet.Cells[row1 + 1, col + i].Style.Font.Bold = true;
                }
                else if (i == 1)
                {
                    worksheet.Cells[row1 + 1, col + i].Value = "Min";
                    worksheet.Cells[row1 + 1, col + i].Style.Font.Color.SetColor(Cores.CorTexto);
                    worksheet.Cells[row1 + 1, col + i].Style.Border.Bottom.Style = ExcelBorderStyle.Thin;
                    worksheet.Cells[row1 + 1, col + i].Style.Font.Bold = true;
                    worksheet.Cells[row1 + 1, col + i].Style.HorizontalAlignment = ExcelHorizontalAlignment.Right;
                }
                else if (i == 2)
                {
                    worksheet.Cells[row1 + 1, col + i].Value = "Q1";
                    worksheet.Cells[row1 + 1, col + i].Style.Font.Color.SetColor(Cores.CorTexto);
                    worksheet.Cells[row1 + 1, col + i].Style.Border.Bottom.Style = ExcelBorderStyle.Thin;
                    worksheet.Cells[row1 + 1, col + i].Style.Font.Bold = true;
                    worksheet.Cells[row1 + 1, col + i].Style.HorizontalAlignment = ExcelHorizontalAlignment.Right;
                }
                else if (i == 3)
                {
                    worksheet.Cells[row1 + 1, col + i].Value = "Median";
                    worksheet.Cells[row1 + 1, col + i].Style.Font.Color.SetColor(Cores.CorTexto);
                    worksheet.Cells[row1 + 1, col + i].Style.Border.Bottom.Style = ExcelBorderStyle.Thin;
                    worksheet.Cells[row1 + 1, col + i].Style.Font.Bold = true;
                    worksheet.Cells[row1 + 1, col + i].Style.HorizontalAlignment = ExcelHorizontalAlignment.Right;
                }
                else if (i == 4)
                {
                    worksheet.Cells[row1 + 1, col + i].Value = "Q3";
                    worksheet.Cells[row1 + 1, col + i].Style.Font.Color.SetColor(Cores.CorTexto);
                    worksheet.Cells[row1 + 1, col + i].Style.Border.Bottom.Style = ExcelBorderStyle.Thin;
                    worksheet.Cells[row1 + 1, col + i].Style.Font.Bold = true;
                    worksheet.Cells[row1 + 1, col + i].Style.HorizontalAlignment = ExcelHorizontalAlignment.Right;
                }
                else if (i == 5)
                {
                    worksheet.Cells[row1 + 1, col + i].Value = "Max";
                    worksheet.Cells[row1 + 1, col + i].Style.Font.Color.SetColor(Cores.CorTexto);
                    worksheet.Cells[row1 + 1, col + i].Style.Border.Bottom.Style = ExcelBorderStyle.Thin;
                    worksheet.Cells[row1 + 1, col + i].Style.Font.Bold = true;
                    worksheet.Cells[row1 + 1, col + i].Style.HorizontalAlignment = ExcelHorizontalAlignment.Right;
                }

            }




            for (int i = 0; i < 6; i++)
            {
                string column = columnName.GetExcelColumnName(col + i);
                int lastRow = peersOutlook.Count() + 5;

                worksheet.Rows[row1 + 2].Height = 3;

                if (i == 0)
                {
                    worksheet.Cells[row1 + 3, col + i].Value = "EV/Book";
                    worksheet.Cells[row1 + 4, col + i].Value = "EV/EBITDA";
                    worksheet.Cells[row1 + 5, col + i].Value = "EV/EBIT";
                    worksheet.Cells[row1 + 6, col + i].Value = "EV/SALES";
                }
                else if (i == 1)
                {

                    worksheet.Cells[row1 + 3, col + i].Formula = "=MIN(G5:G" + lastRow + ")";
                    worksheet.Cells[row1 + 3, col + i].Style.Numberformat.Format = "#,##0.00;(#,##0.00);-";
                    worksheet.Columns[col + i].Width = 12;

                    worksheet.Cells[row1 + 4, col + i].Formula = "=MIN(H5:H" + lastRow + ")";
                    worksheet.Cells[row1 + 4, col + i].Style.Numberformat.Format = "#,##0.00;(#,##0.00);-";
                    worksheet.Columns[col + i].Width = 12;

                    worksheet.Cells[row1 + 5, col + i].Formula = "=MIN(I5:I" + lastRow + ")";
                    worksheet.Cells[row1 + 5, col + i].Style.Numberformat.Format = "#,##0.00;(#,##0.00);-";
                    worksheet.Columns[col + i].Width = 12;

                    worksheet.Cells[row1 + 6, col + i].Formula = "=MIN(J5:J" + lastRow + ")";
                    worksheet.Cells[row1 + 6, col + i].Style.Numberformat.Format = "#,##0.00;(#,##0.00);-";
                    worksheet.Columns[col + i].Width = 12;
                }
                else if (i == 2)
                {

                    worksheet.Cells[row1 + 3, col + i].Formula = "=QUARTILE(G5:G" + lastRow + ",1)";
                    worksheet.Cells[row1 + 3, col + i].Style.Numberformat.Format = "#,##0.00;(#,##0.00);-";
                    worksheet.Columns[col + i].Width = 12;

                    worksheet.Cells[row1 + 4, col + i].Formula = "=QUARTILE(H5:H" + lastRow + ",1)";
                    worksheet.Cells[row1 + 4, col + i].Style.Numberformat.Format = "#,##0.00;(#,##0.00);-";
                    worksheet.Columns[col + i].Width = 12;

                    worksheet.Cells[row1 + 5, col + i].Formula = "=QUARTILE(I5:I" + lastRow + ",1)";
                    worksheet.Cells[row1 + 5, col + i].Style.Numberformat.Format = "#,##0.00;(#,##0.00);-";
                    worksheet.Columns[col + i].Width = 12;

                    worksheet.Cells[row1 + 6, col + i].Formula = "=QUARTILE(J5:J" + lastRow + ",1)";
                    worksheet.Cells[row1 + 6, col + i].Style.Numberformat.Format = "#,##0.00;(#,##0.00);-";
                    worksheet.Columns[col + i].Width = 12;
                }
                else if (i == 3)
                {

                    worksheet.Cells[row1 + 3, col + i].Formula = "=MEDIAN(G5:G" + lastRow + ")";
                    worksheet.Cells[row1 + 3, col + i].Style.Numberformat.Format = "#,##0.00;(#,##0.00);-";
                    worksheet.Columns[col + i].Width = 12;

                    worksheet.Cells[row1 + 4, col + i].Formula = "=MEDIAN(H5:H" + lastRow + ")";
                    worksheet.Cells[row1 + 4, col + i].Style.Numberformat.Format = "#,##0.00;(#,##0.00);-";
                    worksheet.Columns[col + i].Width = 12;

                    worksheet.Cells[row1 + 5, col + i].Formula = "=MEDIAN(I5:I" + lastRow + ")";
                    worksheet.Cells[row1 + 5, col + i].Style.Numberformat.Format = "#,##0.00;(#,##0.00);-";
                    worksheet.Columns[col + i].Width = 12;

                    worksheet.Cells[row1 + 6, col + i].Formula = "=MEDIAN(J5:J" + lastRow + ")";
                    worksheet.Cells[row1 + 6, col + i].Style.Numberformat.Format = "#,##0.00;(#,##0.00);-";
                    worksheet.Columns[col + i].Width = 12;
                }
                else if (i == 4)
                {

                    worksheet.Cells[row1 + 3, col + i].Formula = "=QUARTILE(G5:G" + lastRow + ",3)";
                    worksheet.Cells[row1 + 3, col + i].Style.Numberformat.Format = "#,##0.00;(#,##0.00);-";
                    worksheet.Columns[col + i].Width = 12;

                    worksheet.Cells[row1 + 4, col + i].Formula = "=QUARTILE(H5:H" + lastRow + ",3)";
                    worksheet.Cells[row1 + 4, col + i].Style.Numberformat.Format = "#,##0.00;(#,##0.00);-";
                    worksheet.Columns[col + i].Width = 12;

                    worksheet.Cells[row1 + 5, col + i].Formula = "=QUARTILE(I5:I" + lastRow + ",3)";
                    worksheet.Cells[row1 + 5, col + i].Style.Numberformat.Format = "#,##0.00;(#,##0.00);-";
                    worksheet.Columns[col + i].Width = 12;

                    worksheet.Cells[row1 + 6, col + i].Formula = "=QUARTILE(J5:J" + lastRow + ",3)";
                    worksheet.Cells[row1 + 6, col + i].Style.Numberformat.Format = "#,##0.00;(#,##0.00);-";
                    worksheet.Columns[col + i].Width = 12;
                }
                else if (i == 5)
                {

                    worksheet.Cells[row1 + 3, col + i].Formula = "=MAX(G5:G" + lastRow + ")";
                    worksheet.Cells[row1 + 3, col + i].Style.Numberformat.Format = "#,##0.00;(#,##0.00);-";
                    worksheet.Columns[col + i].Width = 12;

                    worksheet.Cells[row1 + 4, col + i].Formula = "=MAX(H5:H" + lastRow + ")";
                    worksheet.Cells[row1 + 4, col + i].Style.Numberformat.Format = "#,##0.00;(#,##0.00);-";
                    worksheet.Columns[col + i].Width = 12;

                    worksheet.Cells[row1 + 5, col + i].Formula = "=MAX(I5:I" + lastRow + ")";
                    worksheet.Cells[row1 + 5, col + i].Style.Numberformat.Format = "#,##0.00;(#,##0.00);-";
                    worksheet.Columns[col + i].Width = 12;

                    worksheet.Cells[row1 + 6, col + i].Formula = "=MAX(J5:J" + lastRow + ")";
                    worksheet.Cells[row1 + 6, col + i].Style.Numberformat.Format = "#,##0.00;(#,##0.00);-";
                    worksheet.Columns[col + i].Width = 12;
                }

                
            }
            row1 = row1 + 9;


            //Equity Company
            for (int i = 0; i < 6; i++)
            {
                worksheet.Cells[row1, col + i].Style.Fill.PatternType = ExcelFillStyle.Solid;
                worksheet.Cells[row1, col + i].Style.Fill.BackgroundColor.SetColor(Cores.CorPrincipal);

                if (i == 0)
                {
                    worksheet.Cells[row1, col + i].Value = "Company Equity (In Millions)";
                    worksheet.Cells[row1, col + i].Style.Font.Bold = true;
                    worksheet.Cells[row1, col + i].Style.Font.Color.SetColor(Color.White);

                    worksheet.Cells[row1 + 1, col].Value = "Multiples";
                    worksheet.Cells[row1 + 1, col + i].Style.Font.Color.SetColor(Cores.CorTexto);
                    worksheet.Cells[row1 + 1, col + i].Style.Border.Bottom.Style = ExcelBorderStyle.Thin;
                    worksheet.Cells[row1 + 1, col + i].Style.Font.Bold = true;
                }
                else if (i == 1)
                {
                    worksheet.Cells[row1 + 1, col + i].Value = "Min";
                    worksheet.Cells[row1 + 1, col + i].Style.Font.Color.SetColor(Cores.CorTexto);
                    worksheet.Cells[row1 + 1, col + i].Style.Border.Bottom.Style = ExcelBorderStyle.Thin;
                    worksheet.Cells[row1 + 1, col + i].Style.Font.Bold = true;
                    worksheet.Cells[row1 + 1, col + i].Style.HorizontalAlignment = ExcelHorizontalAlignment.Right;
                }
                else if (i == 2)
                {
                    worksheet.Cells[row1 + 1, col + i].Value = "Q1";
                    worksheet.Cells[row1 + 1, col + i].Style.Font.Color.SetColor(Cores.CorTexto);
                    worksheet.Cells[row1 + 1, col + i].Style.Border.Bottom.Style = ExcelBorderStyle.Thin;
                    worksheet.Cells[row1 + 1, col + i].Style.Font.Bold = true;
                    worksheet.Cells[row1 + 1, col + i].Style.HorizontalAlignment = ExcelHorizontalAlignment.Right;
                }
                else if (i == 3)
                {
                    worksheet.Cells[row1 + 1, col + i].Value = "Median";
                    worksheet.Cells[row1 + 1, col + i].Style.Font.Color.SetColor(Cores.CorTexto);
                    worksheet.Cells[row1 + 1, col + i].Style.Border.Bottom.Style = ExcelBorderStyle.Thin;
                    worksheet.Cells[row1 + 1, col + i].Style.Font.Bold = true;
                    worksheet.Cells[row1 + 1, col + i].Style.HorizontalAlignment = ExcelHorizontalAlignment.Right;
                }
                else if (i == 4)
                {
                    worksheet.Cells[row1 + 1, col + i].Value = "Q3";
                    worksheet.Cells[row1 + 1, col + i].Style.Font.Color.SetColor(Cores.CorTexto);
                    worksheet.Cells[row1 + 1, col + i].Style.Border.Bottom.Style = ExcelBorderStyle.Thin;
                    worksheet.Cells[row1 + 1, col + i].Style.Font.Bold = true;
                    worksheet.Cells[row1 + 1, col + i].Style.HorizontalAlignment = ExcelHorizontalAlignment.Right;
                }
                else if (i == 5)
                {
                    worksheet.Cells[row1 + 1, col + i].Value = "Max";
                    worksheet.Cells[row1 + 1, col + i].Style.Font.Color.SetColor(Cores.CorTexto);
                    worksheet.Cells[row1 + 1, col + i].Style.Border.Bottom.Style = ExcelBorderStyle.Thin;
                    worksheet.Cells[row1 + 1, col + i].Style.Font.Bold = true;
                    worksheet.Cells[row1 + 1, col + i].Style.HorizontalAlignment = ExcelHorizontalAlignment.Right;
                }

            }




            for (int i = 0; i < 6; i++)
            {
                string column = columnName.GetExcelColumnName(col + i);
                int lastRow = peersOutlook.Count() + 5;

                if (i==0)
                {
                    row1--;
                }
                

                if (i == 0)
                {
                    worksheet.Cells[row1 + 3, col + i].Value = "EV/Book";
                    worksheet.Cells[row1 + 4, col + i].Value = "EV/EBITDA";
                    worksheet.Cells[row1 + 5, col + i].Value = "EV/EBIT";
                    worksheet.Cells[row1 + 6, col + i].Value = "EV/SALES";
                }
                else if (i == 1)
                {

                    worksheet.Cells[row1 + 3, col + i].Formula = "=C10*"+ column + "5/1000";
                    worksheet.Cells[row1 + 3, col + i].Style.Numberformat.Format = "#,##0.00;(#,##0.00);-";
                    worksheet.Columns[col + i].Width = 12;

                    worksheet.Cells[row1 + 4, col + i].Formula = "=C6*" + column + "6/1000";
                    worksheet.Cells[row1 + 4, col + i].Style.Numberformat.Format = "#,##0.00;(#,##0.00);-";
                    worksheet.Columns[col + i].Width = 12;

                    worksheet.Cells[row1 + 5, col + i].Formula = "=C7*" + column + "7/1000";
                    worksheet.Cells[row1 + 5, col + i].Style.Numberformat.Format = "#,##0.00;(#,##0.00);-";
                    worksheet.Columns[col + i].Width = 12;

                    worksheet.Cells[row1 + 6, col + i].Formula = "=C5*" + column + "8/1000";
                    worksheet.Cells[row1 + 6, col + i].Style.Numberformat.Format = "#,##0.00;(#,##0.00);-";
                    worksheet.Columns[col + i].Width = 12;
                }
                else if (i == 2)
                {

                    worksheet.Cells[row1 + 3, col + i].Formula = "=C10*" + column + "5/1000";
                    worksheet.Cells[row1 + 3, col + i].Style.Numberformat.Format = "#,##0.00;(#,##0.00);-";
                    worksheet.Columns[col + i].Width = 12;

                    worksheet.Cells[row1 + 4, col + i].Formula = "=C6*" + column + "6/1000";
                    worksheet.Cells[row1 + 4, col + i].Style.Numberformat.Format = "#,##0.00;(#,##0.00);-";
                    worksheet.Columns[col + i].Width = 12;

                    worksheet.Cells[row1 + 5, col + i].Formula = "=C7*" + column + "7/1000";
                    worksheet.Cells[row1 + 5, col + i].Style.Numberformat.Format = "#,##0.00;(#,##0.00);-";
                    worksheet.Columns[col + i].Width = 12;

                    worksheet.Cells[row1 + 6, col + i].Formula = "=C5*" + column + "8/1000";
                    worksheet.Cells[row1 + 6, col + i].Style.Numberformat.Format = "#,##0.00;(#,##0.00);-";
                    worksheet.Columns[col + i].Width = 12;
                }
                else if (i == 3)
                {

                    worksheet.Cells[row1 + 3, col + i].Formula = "=C10*" + column + "5/1000";
                    worksheet.Cells[row1 + 3, col + i].Style.Numberformat.Format = "#,##0.00;(#,##0.00);-";
                    worksheet.Columns[col + i].Width = 12;

                    worksheet.Cells[row1 + 4, col + i].Formula = "=C6*" + column + "6/1000";
                    worksheet.Cells[row1 + 4, col + i].Style.Numberformat.Format = "#,##0.00;(#,##0.00);-";
                    worksheet.Columns[col + i].Width = 12;

                    worksheet.Cells[row1 + 5, col + i].Formula = "=C7*" + column + "7/1000";
                    worksheet.Cells[row1 + 5, col + i].Style.Numberformat.Format = "#,##0.00;(#,##0.00);-";
                    worksheet.Columns[col + i].Width = 12;

                    worksheet.Cells[row1 + 6, col + i].Formula = "=C5*" + column + "8/1000";
                    worksheet.Cells[row1 + 6, col + i].Style.Numberformat.Format = "#,##0.00;(#,##0.00);-";
                    worksheet.Columns[col + i].Width = 12;
                }
                else if (i == 4)
                {

                    worksheet.Cells[row1 + 3, col + i].Formula = "=C10*" + column + "5/1000";
                    worksheet.Cells[row1 + 3, col + i].Style.Numberformat.Format = "#,##0.00;(#,##0.00);-";
                    worksheet.Columns[col + i].Width = 12;

                    worksheet.Cells[row1 + 4, col + i].Formula = "=C6*" + column + "6/1000";
                    worksheet.Cells[row1 + 4, col + i].Style.Numberformat.Format = "#,##0.00;(#,##0.00);-";
                    worksheet.Columns[col + i].Width = 12;

                    worksheet.Cells[row1 + 5, col + i].Formula = "=C7*" + column + "7/1000";
                    worksheet.Cells[row1 + 5, col + i].Style.Numberformat.Format = "#,##0.00;(#,##0.00);-";
                    worksheet.Columns[col + i].Width = 12;

                    worksheet.Cells[row1 + 6, col + i].Formula = "=C5*" + column + "8/1000";
                    worksheet.Cells[row1 + 6, col + i].Style.Numberformat.Format = "#,##0.00;(#,##0.00);-";
                    worksheet.Columns[col + i].Width = 12;
                }
                else if (i == 5)
                {

                    worksheet.Cells[row1 + 3, col + i].Formula = "=C10*" + column + "5/1000";
                    worksheet.Cells[row1 + 3, col + i].Style.Numberformat.Format = "#,##0.00;(#,##0.00);-";
                    worksheet.Columns[col + i].Width = 12;

                    worksheet.Cells[row1 + 4, col + i].Formula = "=C6*" + column + "6/1000";
                    worksheet.Cells[row1 + 4, col + i].Style.Numberformat.Format = "#,##0.00;(#,##0.00);-";
                    worksheet.Columns[col + i].Width = 12;

                    worksheet.Cells[row1 + 5, col + i].Formula = "=C7*" + column + "7/1000";
                    worksheet.Cells[row1 + 5, col + i].Style.Numberformat.Format = "#,##0.00;(#,##0.00);-";
                    worksheet.Columns[col + i].Width = 12;

                    worksheet.Cells[row1 + 6, col + i].Formula = "=C5*" + column + "8/1000";
                    worksheet.Cells[row1 + 6, col + i].Style.Numberformat.Format = "#,##0.00;(#,##0.00);-";
                    worksheet.Columns[col + i].Width = 12;
                }
            }

        }
    }
}
