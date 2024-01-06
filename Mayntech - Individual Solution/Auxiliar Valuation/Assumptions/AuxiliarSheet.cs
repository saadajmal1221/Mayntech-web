using Mayntech___Individual_Solution.Auxiliar.Auxiliar.FMP;
using Mayntech___Individual_Solution.Auxiliar.Auxiliar.Other;
using OfficeOpenXml;
using OfficeOpenXml.Style;
using System.Data.Common;
using System.Drawing;


namespace Mayntech___Individual_Solution.Auxiliar_Valuation.Assumptions
{
    public class AuxiliarSheet
    {
        public void AuxiliarSheetConstruction(ExcelPackage package, List<FinancialStatements> balanceSheet, 
            List<FinancialStatements> IncomeStatement, int numberOfYears)
        {
            ExcelNextCol columnName = new ExcelNextCol();
            var worksheet = package.Workbook.Worksheets.Add("Auxiliar");
            worksheet.View.ShowGridLines = false;
            worksheet.View.ZoomScale = 80;
            worksheet.View.FreezePanes(4, 3);
            int row = 2;
            int col = 2;

            List<double> investedCapital = new List<double>();

            for (int i = 0; i < balanceSheet.Count(); i++)
            {
                double PPE = balanceSheet[i].propertyPlantEquipmentNet;
                double Intangibles = balanceSheet[i].IntangibleAssets;
                double NWC = balanceSheet[i].NetReceivables + balanceSheet[i].Inventory + balanceSheet[i].OtherCurrentAssets - (balanceSheet[i].AccountPayables + balanceSheet[i].DeferredRevenue + balanceSheet[i].OtherCurrentLiabilities);

                investedCapital.Add(PPE + Intangibles + NWC);
            }


            for (int i = 0; i < numberOfYears+12; i++)
            {
                if (i<numberOfYears+1)
                {
                    worksheet.Cells[row, col + i].Style.Fill.PatternType = ExcelFillStyle.Solid;
                    worksheet.Cells[row, col + i].Style.Fill.BackgroundColor.SetColor(Cores.CorPrincipal);

                    worksheet.Cells[row + 34, col + 0].Style.Fill.PatternType = ExcelFillStyle.Solid;
                    worksheet.Cells[row + 34, col + 0].Style.Fill.BackgroundColor.SetColor(Cores.CorPrincipal);
                    worksheet.Cells[row + 34, col + 1].Style.Fill.PatternType = ExcelFillStyle.Solid;
                    worksheet.Cells[row + 34, col + 1].Style.Fill.BackgroundColor.SetColor(Cores.CorPrincipal);
                }


                worksheet.Cells[row + 9, col + i].Style.Fill.PatternType = ExcelFillStyle.Solid;
                worksheet.Cells[row + 9, col + i].Style.Fill.BackgroundColor.SetColor(Cores.CorPrincipal);

                worksheet.Cells[row + 15, col + i].Style.Fill.PatternType = ExcelFillStyle.Solid;
                worksheet.Cells[row + 15, col + i].Style.Fill.BackgroundColor.SetColor(Cores.CorPrincipal);

                //worksheet.Cells[row + 20, col + i].Style.Fill.PatternType = ExcelFillStyle.Solid;
                //worksheet.Cells[row + 20, col + i].Style.Fill.BackgroundColor.SetColor(Cores.CorPrincipal);




                if (i==0)
                {
                    worksheet.Cells[row, col].Value = "Cash Conversion cycle";
                    worksheet.Cells[row, col].Style.Font.Bold = true;
                    worksheet.Cells[row, col].Style.Font.Color.SetColor(Color.White);

                    worksheet.Cells[row + 1, col + i].Value = "Description";
                    worksheet.Cells[row + 1, col + i].Style.Font.Color.SetColor(Cores.CorTexto);
                    worksheet.Cells[row + 1, col + i].Style.Border.Bottom.Style = ExcelBorderStyle.Thin;
                    worksheet.Cells[row + 1, col + i].Style.Font.Bold = true;

                    worksheet.Cells[row + 3, col + i].Value = "DPO";
                    worksheet.Cells[row + 4, col + i].Value = "DSO";
                    worksheet.Cells[row + 5, col + i].Value = "DIO";
                    worksheet.Cells[row + 6, col + i].Value = "CCC";




                    worksheet.Cells[row+9, col].Value = "Depreciation And Amortization Rate";
                    worksheet.Cells[row+9, col].Style.Font.Bold = true;
                    worksheet.Cells[row + 9, col].Style.Font.Color.SetColor(Color.White);

                    worksheet.Cells[row + 10, col + i].Value = "Description";
                    worksheet.Cells[row + 10, col + i].Style.Font.Color.SetColor(Cores.CorTexto);
                    worksheet.Cells[row + 10, col + i].Style.Border.Bottom.Style = ExcelBorderStyle.Thin;
                    worksheet.Cells[row + 10, col + i].Style.Font.Bold = true;
                    worksheet.Cells[row + 11, col + i].Value = "D&A";
                    worksheet.Cells[row + 12, col + i].Value = "D&A rate";
                    worksheet.Cells[row + 12, col + i].Style.Font.Color.SetColor(Cores.corSecundáriaText);



                    worksheet.Cells[row + 15, col].Value = "Invested Capital";
                    worksheet.Cells[row + 15, col].Style.Font.Bold = true;
                    worksheet.Cells[row + 15, col].Style.Font.Color.SetColor(Color.White);

                    worksheet.Cells[row + 16, col + i].Value = "Description";
                    worksheet.Cells[row + 16, col + i].Style.Font.Color.SetColor(Cores.CorTexto);
                    worksheet.Cells[row + 16, col + i].Style.Border.Bottom.Style = ExcelBorderStyle.Thin;
                    worksheet.Cells[row + 16, col + i].Style.Font.Bold = true;


                    worksheet.Cells[row + 17, col + i].Value = "PP&E";
                    worksheet.Cells[row + 18, col + i].Value = "% Change";
                    worksheet.Cells[row + 18, col + i].Style.Font.Color.SetColor(Cores.corSecundáriaText);

                    worksheet.Cells[row + 20, col + i].Value = "Intangibles";
                    worksheet.Cells[row + 21, col + i].Value = "% Change";
                    worksheet.Cells[row + 21, col + i].Style.Font.Color.SetColor(Cores.corSecundáriaText);

                    worksheet.Cells[row + 23, col + i].Value = "Goodwill";
                    worksheet.Cells[row + 24, col + i].Value = "% Change";
                    worksheet.Cells[row + 24, col + i].Style.Font.Color.SetColor(Cores.corSecundáriaText);

                    worksheet.Rows[row + 23].OutlineLevel = 1;
                    worksheet.Rows[row + 23].Collapsed = true;
                    worksheet.Rows[row + 24].OutlineLevel = 1;
                    worksheet.Rows[row + 24].Collapsed = true;
                    worksheet.Rows[row + 25].OutlineLevel = 1;
                    worksheet.Rows[row + 25].Collapsed = true;

                    worksheet.Cells[row + 26, col + i].Value = "NWC";
                    worksheet.Cells[row + 27, col + i].Value = "% revenues";
                    worksheet.Cells[row + 27, col + i].Style.Font.Color.SetColor(Cores.corSecundáriaText);

                    worksheet.Cells[row + 29, col + i].Value = "Invested Capital";
                    worksheet.Cells[row + 29, col + i].Style.Fill.PatternType = ExcelFillStyle.Solid;
                    worksheet.Cells[row + 29, col + i].Style.Fill.BackgroundColor.SetColor(Cores.corSecundária);
                    worksheet.Cells[row + 29, col + i].Style.Font.Bold = true;

                    worksheet.Cells[row + 31, col + i].Value = "Net Investment";
                    worksheet.Cells[row + 31, col + i].Style.Fill.PatternType = ExcelFillStyle.Solid;
                    worksheet.Cells[row + 31, col + i].Style.Fill.BackgroundColor.SetColor(Cores.CorCinzentoBackground);
                    worksheet.Cells[row + 31, col + i].Style.Font.Bold = true;


                    worksheet.Cells[row + 34, col].Value = "Capital Structure";
                    worksheet.Cells[row + 34, col].Style.Font.Bold = true;
                    worksheet.Cells[row + 34, col].Style.Font.Color.SetColor(Color.White);

                    worksheet.Cells[row + 35, col + i].Value = "Description";
                    worksheet.Cells[row + 35, col + i].Style.Font.Color.SetColor(Cores.CorTexto);
                    worksheet.Cells[row + 35, col + i].Style.Border.Bottom.Style = ExcelBorderStyle.Thin;
                    worksheet.Cells[row + 35, col + i].Style.Font.Bold = true;

                    worksheet.Cells[row + 37, col + i].Value = "Debt to Capital";


                    //worksheet.Cells[row + 20, col].Value = "Operating Working Capital";
                    //worksheet.Cells[row + 20, col].Style.Font.Bold = true;
                    //worksheet.Cells[row + 20, col].Style.Font.Color.SetColor(Color.White);

                    //worksheet.Cells[row + 21, col + i].Value = "Description";
                    //worksheet.Cells[row + 21, col + i].Style.Font.Color.SetColor(Cores.CorTexto);
                    //worksheet.Cells[row + 21, col + i].Style.Border.Bottom.Style = ExcelBorderStyle.Thin;
                    //worksheet.Cells[row + 21, col + i].Style.Font.Bold = true;
                    //worksheet.Cells[row + 22, col + i].Value = "Operating NWC";
                    //worksheet.Cells[row + 23, col + i].Value = "Change in NWC";

                    worksheet.Columns[col + i].Width = 25;
                }
                else if (i < numberOfYears + 1)
                {
                    string column = columnName.GetExcelColumnName(col + i);
                    string columnLeft = columnName.GetExcelColumnName(col + i-1);
                    worksheet.Cells[row + 1, col + i].Formula = "=YEAR('P&L'!" + column + "3)";
                    worksheet.Cells[row + 1, col + i].Style.Font.Color.SetColor(Cores.CorTexto);
                    worksheet.Cells[row + 1, col + i].Style.Border.Bottom.Style = ExcelBorderStyle.Thin;
                    worksheet.Cells[row + 1, col + i].Style.Font.Bold = true;

                    worksheet.Rows[row + 2].Height = 3;

                    worksheet.Cells[row + 3, col + i].Formula = "=('BS'!" + column + "27/-'P&L'!" + column + "6)*365";
                    worksheet.Cells[row + 3, col + i].Style.Numberformat.Format = "0";
                    worksheet.Cells[row + 4, col + i].Formula = "=('BS'!" + column + "9/'P&L'!" + column + "5)*365";
                    worksheet.Cells[row + 4, col + i].Style.Numberformat.Format = "0";
                    worksheet.Cells[row + 5, col + i].Formula = "=('BS'!" + column + "10/-'P&L'!" + column + "6)*365";
                    worksheet.Cells[row + 5, col + i].Style.Numberformat.Format = "0";
                    worksheet.Cells[row + 6, col + i].Formula = "=-" +column + "5+" +column + "6+" + column + "7";
                    worksheet.Cells[row + 6, col + i].Style.Numberformat.Format = "0";



                    //DP&A
                    worksheet.Cells[row + 10, col + i].Formula = "=YEAR('P&L'!" + column + "3)";
                    worksheet.Cells[row + 10, col + i].Style.Font.Color.SetColor(Cores.CorTexto);
                    worksheet.Cells[row + 10, col + i].Style.Border.Bottom.Style = ExcelBorderStyle.Thin;
                    worksheet.Cells[row + 10, col + i].Style.Font.Bold = true;
                    worksheet.Cells[row + 11, col + i].Formula = "=-'P&L'!" + column + "32";
                    worksheet.Cells[row + 11, col + i].Style.Numberformat.Format = "#,##0;(#,##0);-";
                    worksheet.Cells[row + 12, col + i].Formula = "=-'P&L'!" + column + "32/('BS'!" + column+"15+'BS'!" + column + "17)";
                    worksheet.Cells[row + 12, col + i].Style.Numberformat.Format = "0.00%";
                    worksheet.Cells[row + 12, col + i].Style.Font.Color.SetColor(Cores.corSecundáriaText);


                    //Invested Capital

                    worksheet.Cells[row + 16, col + i].Formula = "=YEAR('P&L'!" + column + "3)";
                    worksheet.Cells[row + 16, col + i].Style.Font.Color.SetColor(Cores.CorTexto);
                    worksheet.Cells[row + 16, col + i].Style.Border.Bottom.Style = ExcelBorderStyle.Thin;
                    worksheet.Cells[row + 16, col + i].Style.Font.Bold = true;

                    //PP&E

                    worksheet.Cells[row + 17, col + i].Formula = "='BS'!" + column + "15";
                    worksheet.Cells[row + 17, col + i].Style.Numberformat.Format = "#,##0;(#,##0);-";

                    if (i>1)
                    {
                        worksheet.Cells[row + 18, col + i].Formula = "=(" + column + "19/" + columnLeft + "19)-1";
                        worksheet.Cells[row + 18, col + i].Style.Numberformat.Format = "0.00%";
                        worksheet.Cells[row + 18, col + i].Style.Font.Color.SetColor(Cores.corSecundáriaText);
                    }


                    //Intangibles

                    worksheet.Cells[row + 20, col + i].Formula = "='BS'!" + column + "17";
                    worksheet.Cells[row + 20, col + i].Style.Numberformat.Format = "#,##0;(#,##0);-";

                    if (i > 1)
                    {
                        worksheet.Cells[row + 21, col + i].Formula = "=(" + column + "22/" + columnLeft + "22)-1";
                        worksheet.Cells[row + 21, col + i].Style.Numberformat.Format = "0.00%";
                        worksheet.Cells[row + 21, col + i].Style.Font.Color.SetColor(Cores.corSecundáriaText);
                    }


                    //Goodwill

                    worksheet.Cells[row + 23, col + i].Formula = "='BS'!" + column + "16";
                    worksheet.Cells[row + 23, col + i].Style.Numberformat.Format = "#,##0;(#,##0);-";

                    if (i > 1)
                    {
                        worksheet.Cells[row + 24, col + i].Formula = "=(" + column + "25/" + columnLeft + "25) - 1";
                        worksheet.Cells[row + 24, col + i].Style.Numberformat.Format = "0.00%";
                        worksheet.Cells[row + 24, col + i].Style.Font.Color.SetColor(Cores.corSecundáriaText);
                    }



                    //NWC
                    if (investedCapital.Min() < 0)
                    {
                        worksheet.Cells[row + 26, col + i].Formula = "=('BS'!" + column + "9 + 'BS'!" + column + "10)-('BS'!" + column + "27 + 'BS'!" + column + "30)";
                    }
                    else
                    {
                        worksheet.Cells[row + 26, col + i].Formula = "=('BS'!" + column + "9 + 'BS'!" + column + "10 +  'BS'!" + column + "11)-('BS'!" + column + "27 + 'BS'!" + column + "30+ 'BS'!" + column + "31)";
                    }
                    
                    worksheet.Cells[row + 26, col + i].Style.Numberformat.Format = "#,##0;(#,##0);-";

                    worksheet.Cells[row + 27, col + i].Formula = "=" + column + "28/'P&L'!" + column + "5";
                    worksheet.Cells[row + 27, col + i].Style.Numberformat.Format = "0.00%";
                    worksheet.Cells[row + 27, col + i].Style.Font.Color.SetColor(Cores.corSecundáriaText);

                    worksheet.Columns[col + i].Width = 12;


                    //Operational Invested capital
                    worksheet.Cells[row + 29, col + i].Formula = "="+column + "19 +"+ column + "22 +" + column + "28";
                    worksheet.Cells[row + 29, col + i].Style.Numberformat.Format = "#,##0;(#,##0);-";
                    worksheet.Cells[row + 29, col + i].Style.Fill.PatternType = ExcelFillStyle.Solid;
                    worksheet.Cells[row + 29, col + i].Style.Fill.BackgroundColor.SetColor(Cores.corSecundária);
                    worksheet.Cells[row + 29, col + i].Style.Font.Bold = true;

                    worksheet.Cells[row + 31, col + i].Style.Fill.PatternType = ExcelFillStyle.Solid;
                    worksheet.Cells[row + 31, col + i].Style.Fill.BackgroundColor.SetColor(Cores.CorCinzentoBackground);
                    if (i > 1)
                    {
                        worksheet.Cells[row + 31, col + i].Formula = "=" + column + "31-" + columnLeft + "31";
                        worksheet.Cells[row + 31, col + i].Style.Numberformat.Format = "#,##0;(#,##0);-";
                        worksheet.Cells[row + 31, col + i].Style.Font.Bold = true;
                    }

                    if (i == numberOfYears)
                    {
                        //Capital Structure
                        worksheet.Cells[row + 35, col + 1].Formula = "=YEAR('P&L'!" + column + "3)";
                        worksheet.Cells[row + 35, col + 1].Style.Font.Color.SetColor(Cores.CorTexto);
                        worksheet.Cells[row + 35, col + 1].Style.Border.Bottom.Style = ExcelBorderStyle.Thin;
                        worksheet.Cells[row + 35, col + 1].Style.Font.Bold = true;

                        worksheet.Rows[row + 36].Height = 3;

                        worksheet.Cells[row + 37, col + 1].Formula = "='BS'!" + column + "42/('BS'!" + column + "42 + ('DCF Valuation'!C14*'DCF Valuation'!C15))";
                        worksheet.Cells[row + 37, col + 1].Style.Numberformat.Format = "#,##0.00;(#,##0.00);-";
                    }

                }
                else 
                {
                    string column = columnName.GetExcelColumnName(col + i);
                    string columnLeft = columnName.GetExcelColumnName(col + i - 1);

                    //DP&A
                    worksheet.Cells[row + 10, col + i].Formula = "=" + columnLeft +"12 + 1" ;
                    worksheet.Cells[row + 10, col + i].Style.Font.Color.SetColor(Cores.CorTexto);
                    worksheet.Cells[row + 10, col + i].Style.Border.Bottom.Style = ExcelBorderStyle.Thin;
                    worksheet.Cells[row + 10, col + i].Style.Font.Bold = true;
                    worksheet.Cells[row + 11, col + i].Formula = "=" + column + "14*(" + column + "19 + " + column + "22)" ;
                    worksheet.Cells[row + 11, col + i].Style.Numberformat.Format = "#,##0;(#,##0);-";
                    worksheet.Cells[row + 12, col + i].Formula = "='Assumptions'!H8";
                    worksheet.Cells[row + 12, col + i].Style.Numberformat.Format = "0.00%";
                    worksheet.Cells[row + 12, col + i].Style.Font.Color.SetColor(Cores.corSecundáriaText);


                    string ppe;
                    string intangible;
                    string nwc;

                    if (i < numberOfYears + 1)
                    {
                        ppe = "='Assumptions'!C8";
                        intangible = "='Assumptions'!C9";
                        nwc = "'Assumptions'!C10";
                    }
                    else if (i < numberOfYears + 6)
                    {
                        ppe = "='Assumptions'!D8";
                        intangible = "='Assumptions'!d9";
                        nwc = "'Assumptions'!D10";
                    }
                    else if (i<numberOfYears+10)
                    {
                        ppe = "='Assumptions'!E8";
                        intangible = "='Assumptions'!E9";
                        nwc = "'Assumptions'!E10";
                    }
                    else if(i < numberOfYears + 11)
                    {
                        ppe = "=(" + columnLeft +"20" + "+'Assumptions'!M5)/2";
                        intangible = "=(" + columnLeft + "23" + "+'Assumptions'!M5)/2";
                        nwc = column + "28/'FCF Projections'!" + column + "5";
                    }
                    else 
                    {
                        ppe = "='Assumptions'!M5";
                        intangible = "='Assumptions'!M5";
                        nwc = column + "28/'FCF Projections'!" + column + "5";
                    }

                    //Invested Capital

                    worksheet.Cells[row + 16, col + i].Formula = "=" + columnLeft + "18 + 1";
                    worksheet.Cells[row + 16, col + i].Style.Font.Color.SetColor(Cores.CorTexto);
                    worksheet.Cells[row + 16, col + i].Style.Border.Bottom.Style = ExcelBorderStyle.Thin;
                    worksheet.Cells[row + 16, col + i].Style.Font.Bold = true;

                    //PP&E

                    worksheet.Cells[row + 17, col + i].Formula = "=(1+" + column + "20)*" + columnLeft + "19";
                    worksheet.Cells[row + 17, col + i].Style.Numberformat.Format = "#,##0;(#,##0);-";

                    if (i > 1)
                    {
                        worksheet.Cells[row + 18, col + i].Formula = ppe;
                        worksheet.Cells[row + 18, col + i].Style.Numberformat.Format = "0.00%";
                        worksheet.Cells[row + 18, col + i].Style.Font.Color.SetColor(Cores.corSecundáriaText);
                    }


                    //Intangibles

                    worksheet.Cells[row + 20, col + i].Formula = "=(1+" + column + "23)*" + columnLeft + "22";
                    worksheet.Cells[row + 20, col + i].Style.Numberformat.Format = "#,##0;(#,##0);-";

                    if (i > 1)
                    {
                        worksheet.Cells[row + 21, col + i].Formula = intangible;
                        worksheet.Cells[row + 21, col + i].Style.Numberformat.Format = "0.00%";
                        worksheet.Cells[row + 21, col + i].Style.Font.Color.SetColor(Cores.corSecundáriaText);
                    }


                    //Goodwill

                    worksheet.Cells[row + 23, col + i].Formula = "=(1+" + column + "26)*" + columnLeft + "25";
                    worksheet.Cells[row + 23, col + i].Style.Numberformat.Format = "#,##0;(#,##0);-";

                    if (i > 1)
                    {
                        worksheet.Cells[row + 24, col + i].Value = 0;
                        worksheet.Cells[row + 24, col + i].Style.Numberformat.Format = "0.00%";
                        worksheet.Cells[row + 24, col + i].Style.Font.Color.SetColor(Cores.corSecundáriaText);
                    }



                    //NWC

                    worksheet.Cells[row + 26, col + i].Formula = "=" + columnLeft + "28";
                    worksheet.Cells[row + 26, col + i].Style.Numberformat.Format = "#,##0;(#,##0);-";
                    worksheet.Cells[row + 27, col + i].Formula = "=" + columnLeft + "29" + "*(1+" + nwc +")";
                    worksheet.Cells[row + 27, col + i].Style.Numberformat.Format = "0.00%";
                    worksheet.Cells[row + 27, col + i].Style.Font.Color.SetColor(Cores.corSecundáriaText);

                    worksheet.Columns[col + i].Width = 12;


                    //Operational Invested capital
                    worksheet.Cells[row + 29, col + i].Formula = "=" + column + "19 +" + column + "22 +" + column + "28";
                    worksheet.Cells[row + 29, col + i].Style.Numberformat.Format = "#,##0;(#,##0);-";
                    worksheet.Cells[row + 29, col + i].Style.Fill.PatternType = ExcelFillStyle.Solid;
                    worksheet.Cells[row + 29, col + i].Style.Fill.BackgroundColor.SetColor(Cores.corSecundária);
                    worksheet.Cells[row + 29, col + i].Style.Font.Bold = true;

                    //Net investment
                    worksheet.Cells[row + 31, col + i].Formula = "=" + column + "31-" + columnLeft + "31";
                    worksheet.Cells[row + 31, col + i].Style.Numberformat.Format = "#,##0;(#,##0);-";
                    worksheet.Cells[row + 31, col + i].Style.Fill.PatternType = ExcelFillStyle.Solid;
                    worksheet.Cells[row + 31, col + i].Style.Fill.BackgroundColor.SetColor(Cores.CorCinzentoBackground);
                    worksheet.Cells[row + 31, col + i].Style.Font.Bold = true;

                }

            }
                    


            
        }
    }
}
