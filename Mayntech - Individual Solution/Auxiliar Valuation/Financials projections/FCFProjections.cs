using Mayntech___Individual_Solution.Auxiliar.Auxiliar.Other;
using Mayntech___Individual_Solution.Pages.Solutions;
using Newtonsoft.Json;
using OfficeOpenXml;
using OfficeOpenXml.Style;
using System.Collections.Generic;
using System.Data.Common;
using System.Drawing;
using System.Globalization;
using System.Reflection;
using System.Reflection.Metadata.Ecma335;

namespace Mayntech___Individual_Solution.Auxiliar_Valuation.Financials_projections
{
    public class FCFProjections
    {

        public void IncomeProjectionConstruction(ExcelPackage package,
            int numberOfYears, Taxes tax, string calendarYear)
        {

            int LastYear = int.Parse(calendarYear);
            ExcelNextCol columnLetter = new ExcelNextCol();
            var worksheet = package.Workbook.Worksheets.Add("FCF Projections");
            worksheet.View.ShowGridLines = false;
            worksheet.View.ZoomScale = 80;
            worksheet.View.FreezePanes(4, 3);
            int row = 2;
            int col = 2;


            for (int i = 0; i < numberOfYears+12; i++)
            {
                worksheet.Cells[row, col + i].Style.Fill.PatternType = ExcelFillStyle.Solid;
                worksheet.Cells[row, col + i].Style.Fill.BackgroundColor.SetColor(Cores.CorPrincipal);

                for (int a = 0; a < 10; a++)
                {
                    worksheet.Cells[row + 2 + a, col + i].Style.Fill.PatternType = ExcelFillStyle.Solid;
                    worksheet.Cells[row + 2 + a, col + i].Style.Fill.BackgroundColor.SetColor(Cores.CorCinzentoClaroBackground);

                }

                for (int a = 19; a < 25; a++)
                {
                    worksheet.Cells[row + 2 + a, col + i].Style.Fill.PatternType = ExcelFillStyle.Solid;
                    worksheet.Cells[row + 2 + a, col + i].Style.Fill.BackgroundColor.SetColor(Cores.CorCinzentoClaroBackground);

                }

                if (i<1)
                {
                    worksheet.Cells[row , col + i].Value = "Free cash flow Projection";
                    worksheet.Cells[row, col + i].Style.Font.Bold = true;
                    worksheet.Cells[row, col + i].Style.Font.Color.SetColor(Color.White);

                    worksheet.Cells[row + 1, col + i].Formula = "='P&L'!B3";
                    worksheet.Cells[row + 1, col + i].Style.Font.Bold = true;
                    worksheet.Cells[row + 1, col + i].Style.Border.Bottom.Style = ExcelBorderStyle.Thin;
                    

                    worksheet.Rows[row + 2].Height = 3;
                    worksheet.Rows[row + 20].Height = 2;


                    //Revenue
                    worksheet.Cells[row + 3, col + i].Formula = "='P&L'!B5";
                    worksheet.Cells[row + 4, col + i].Value = "% change";
                    worksheet.Cells[row + 4, col + i].Style.Font.Color.SetColor(Cores.corSecundáriaText);
                    worksheet.Cells[row + 4, col + i].Style.Font.Italic = true;

                    worksheet.Columns[col + i].Width = 35;

                    //COGS
                    AuxHistoricalDescription(worksheet, row + 3, col, i, 6);

                    //OperatingCosts
                    AuxHistoricalDescription(worksheet, row + 6, col, i, 10);

                    //Core EBIT
                    worksheet.Cells[row + 12, col + i].Value = "Core EBIT";
                    worksheet.Cells[row + 12, col + i].Style.Font.Bold = true;
                    worksheet.Cells[row + 12, col + i].Style.Fill.PatternType = ExcelFillStyle.Solid;
                    worksheet.Cells[row + 12, col + i].Style.Fill.BackgroundColor.SetColor(Cores.CorCinzentoBackground);

                    //Less Taxes
                    worksheet.Cells[row + 15, col + i].Value = "Taxes";
                    worksheet.Cells[row + 16, col + i].Value = "Tax rate";
                    worksheet.Cells[row + 16, col + i].Style.Font.Color.SetColor(Cores.corSecundáriaText);

                    worksheet.Cells[row + 16, col + i].Style.Font.Italic = true;
                    
                    worksheet.Rows[row + 14].OutlineLevel = 1;
                    worksheet.Rows[row + 14].Collapsed = true;
                    worksheet.Rows[row + 15].OutlineLevel = 1;
                    worksheet.Rows[row + 15].Collapsed = true;
                    worksheet.Rows[row + 16].OutlineLevel = 1;
                    worksheet.Rows[row + 16].Collapsed = true;
                    worksheet.Rows[row + 17].OutlineLevel = 1;
                    worksheet.Rows[row + 17].Collapsed = true;

                    //Earnings before interest adjusted for taxes
                    worksheet.Cells[row + 18, col + i].Value = "NOPLAT";
                    worksheet.Cells[row + 18, col + i].Style.Fill.PatternType = ExcelFillStyle.Solid;
                    worksheet.Cells[row + 18, col + i].Style.Fill.BackgroundColor.SetColor(Cores.corSecundária);
                    worksheet.Cells[row + 18, col + i].Style.Font.Bold = true;

                    ////Capex
                    //worksheet.Cells[row + 21, col + i].Value = "Capex";
                    //worksheet.Cells[row + 22, col + i].Value = "% change";
                    //worksheet.Cells[row + 22, col + i].Style.Font.Color.SetColor(Cores.corSecundáriaText);
                    //worksheet.Cells[row + 22, col + i].Style.Font.Italic = true;


                    ////Intangible assets
                    //worksheet.Cells[row + 24, col + i].Value = "Change in Intangibles";
                    //worksheet.Cells[row + 25, col + i].Value = "% change";
                    //worksheet.Cells[row + 25, col + i].Style.Font.Color.SetColor(Cores.corSecundáriaText);
                    //worksheet.Cells[row + 25, col + i].Style.Font.Italic = true;

                    //Invested Capital
                    worksheet.Cells[row + 21, col + i].Value = "Net Investment";
                    worksheet.Cells[row + 22, col + i].Value = "% change";
                    worksheet.Cells[row + 22, col + i].Style.Font.Color.SetColor(Cores.corSecundáriaText);
                    worksheet.Cells[row + 22, col + i].Style.Font.Italic = true;


                    //DP&A
                    worksheet.Cells[row + 24, col + i].Value = "D&A";
                    worksheet.Cells[row + 25, col + i].Value = "D&A rate";
                    worksheet.Cells[row + 25, col + i].Style.Font.Color.SetColor(Cores.corSecundáriaText);
                    worksheet.Cells[row + 25, col + i].Style.Font.Italic = true;

                    ////NWC
                    //worksheet.Cells[row + 30, col + i].Value = "Change in NWC";
                    //worksheet.Cells[row + 31, col + i].Value = "% change";
                    //worksheet.Cells[row + 31, col + i].Style.Font.Color.SetColor(Cores.corSecundáriaText);
                    //worksheet.Cells[row + 31, col + i].Style.Font.Italic = true;




                    //FCF
                    worksheet.Cells[row + 27, col + i].Value = "Free Cash Flow";
                    worksheet.Cells[row + 27, col + i].Style.Fill.PatternType = ExcelFillStyle.Solid;
                    worksheet.Cells[row + 27, col + i].Style.Fill.BackgroundColor.SetColor(Cores.corSecundária);
                    worksheet.Cells[row + 27, col + i].Style.Font.Bold = true;

                }
                else if (i<numberOfYears+1)
                {


                    string columnLeft = columnLetter.GetExcelColumnName(col + i-1);
                    string column = columnLetter.GetExcelColumnName(col + i);
                    worksheet.Cells[row + 1, col + i].Formula = "=YEAR('P&L'!" +column + "3)";
                    worksheet.Cells[row + 1, col + i].Style.Font.Bold = true;
                    worksheet.Cells[row + 1, col + i].Style.Border.Bottom.Style = ExcelBorderStyle.Thin;
                    worksheet.Columns[col + i].Width = 12;

                    //Revenue
                    worksheet.Cells[row + 3, col + i].Formula = "='P&L'!" + column + "5";
                    worksheet.Cells[row + 3, col + i].Style.Numberformat.Format = "#,##0;(#,##0);-";


                    //Gross Profit
                    AuxHistoricalValues(worksheet, row + 3, col, i, 6, "N");

                    //Operating Costs
                    AuxHistoricalValues(worksheet, row + 6, col, i, 10, "N");

                    //Core EBIT
                    int rowRevenue = row + 3;
                    int rowCostOfRevenue = row + 6;
                    int rowOpCosts = row + 9;
                    int rowEBIT = row + 12;
                    int rowTaxeRate = row + 16;
                    int rowTaxes = row + 15;
                    int rowEBIAT = row + 18;
                    int rowInvestedCapital = row + 21;
                    int rowDPA = row + 24;


                    worksheet.Cells[row + 12, col + i].Formula = "=" + column + rowRevenue + "-" + column + rowOpCosts + "-" + column + rowCostOfRevenue;
                    worksheet.Cells[row + 12, col + i].Style.Numberformat.Format = "#,##0;(#,##0);-";
                    worksheet.Cells[row + 12, col + i].Style.Font.Bold = true;
                    worksheet.Cells[row + 12, col + i].Style.Fill.PatternType = ExcelFillStyle.Solid;
                    worksheet.Cells[row + 12, col + i].Style.Fill.BackgroundColor.SetColor(Cores.CorCinzentoBackground);

                    //Less Taxes
                    worksheet.Cells[row + 15, col + i].Formula = "=" + column + rowEBIT + "*" + column + rowTaxeRate;
                    worksheet.Cells[row + 15, col + i].Style.Numberformat.Format = "#,##0;(#,##0);-";

                    worksheet.Cells[row + 16, col + i].Value = GetTaxByYear(tax, LastYear, numberOfYears, i) / 100;
                    
                    worksheet.Cells[row + 16, col + i].Style.Numberformat.Format = "0.00%";
                    worksheet.Cells[row + 16, col + i].Style.Font.Color.SetColor(Cores.corSecundáriaText);
                    worksheet.Cells[row + 16, col + i].Style.Font.Italic = true;

                    //Earnings before interest adjusted for taxes
                    worksheet.Cells[row + 18, col + i].Formula = "=" + column + rowEBIT + "-" + column + rowTaxes;
                    worksheet.Cells[row + 18, col + i].Style.Numberformat.Format = "#,##0;(#,##0);-";
                    worksheet.Cells[row + 18, col + i].Style.Fill.PatternType = ExcelFillStyle.Solid;
                    worksheet.Cells[row + 18, col + i].Style.Fill.BackgroundColor.SetColor(Cores.corSecundária);
                    worksheet.Cells[row + 18, col + i].Style.Font.Bold = true;

                    ////Capex
                    //worksheet.Cells[row + 21, col + i].Formula = "='Auxiliar'!" +column + "18";
                    //worksheet.Cells[row + 21, col + i].Style.Numberformat.Format = "#,##0;(#,##0);-";


                    ////Intangible assets
                    //worksheet.Cells[row + 24, col + i].Formula = "='Auxiliar'!" + column + "19";
                    //worksheet.Cells[row + 24, col + i].Style.Numberformat.Format = "#,##0;(#,##0);-";

                    //Net Investment
                    worksheet.Cells[row + 21, col + i].Formula = "='Auxiliar'!" + column + "33";
                    worksheet.Cells[row + 21, col + i].Style.Numberformat.Format = "#,##0;(#,##0);-";

                    worksheet.Cells[row + 22, col + i].Formula = "=(" + column + "23/" + columnLeft + "23)-1";
                    worksheet.Cells[row + 22, col + i].Style.Font.Color.SetColor(Cores.corSecundáriaText);
                    worksheet.Cells[row + 22, col + i].Style.Font.Italic = true;
                    worksheet.Cells[row + 22, col + i].Style.Numberformat.Format = "0.00%";

                    //DP&A
                    worksheet.Cells[row + 24, col + i].Formula = "='Auxiliar'!" + column + "13";
                    worksheet.Cells[row + 24, col + i].Style.Numberformat.Format = "#,##0;(#,##0);-";

                    worksheet.Cells[row + 25, col + i].Formula = "='Auxiliar'!" + column + "14";
                    worksheet.Cells[row + 25, col + i].Style.Font.Color.SetColor(Cores.corSecundáriaText);
                    worksheet.Cells[row + 25, col + i].Style.Font.Italic = true;
                    worksheet.Cells[row + 25, col + i].Style.Numberformat.Format = "0.00%";

                    ////NWC
                    //worksheet.Cells[row + 30, col + i].Formula = "='Auxiliar'!" + column + "25";
                    //worksheet.Cells[row + 30, col + i].Style.Numberformat.Format = "#,##0;(#,##0);-";

                    //FCF
                    worksheet.Cells[row + 27, col + i].Formula = "=" + column + rowEBIAT + "-" + column + rowInvestedCapital + "+" + column + rowDPA;
                    worksheet.Cells[row + 27, col + i].Style.Numberformat.Format = "#,##0;(#,##0);-";
                    worksheet.Cells[row + 27, col + i].Style.Fill.PatternType = ExcelFillStyle.Solid;
                    worksheet.Cells[row + 27, col + i].Style.Fill.BackgroundColor.SetColor(Cores.corSecundária);
                    worksheet.Cells[row + 27, col + i].Style.Font.Bold = true;

                    if (i>1)
                    {
                        
                        worksheet.Cells[row + 4, col + i].Formula = "=(" + column + "5/"+columnLeft + "5)-1";
                        worksheet.Cells[row + 4, col + i].Style.Font.Color.SetColor(Cores.corSecundáriaText);
                        worksheet.Cells[row + 4, col + i].Style.Font.Italic = true;
                        worksheet.Cells[row + 4, col + i].Style.Numberformat.Format = "0.00%";


                        ////Capex
                        //worksheet.Cells[row + 22, col + i].Formula = "=(" + column + "23/" + columnLeft + "23)-1";
                        //worksheet.Cells[row + 22, col + i].Style.Font.Color.SetColor(Cores.corSecundáriaText);
                        //worksheet.Cells[row + 22, col + i].Style.Font.Italic = true;
                        //worksheet.Cells[row + 22, col + i].Style.Numberformat.Format = "0.00%";

                        ////Intangible
                        //worksheet.Cells[row + 25, col + i].Formula = "=(" + column + "26/" + columnLeft + "26)-1";
                        //worksheet.Cells[row + 25, col + i].Style.Font.Color.SetColor(Cores.corSecundáriaText);
                        //worksheet.Cells[row + 25, col + i].Style.Font.Italic = true;
                        //worksheet.Cells[row + 25, col + i].Style.Numberformat.Format = "0.00%";




                        ////NWC
                        //worksheet.Cells[row + 31, col + i].Formula = "=(" + column + "32/" + columnLeft + "32)-1";
                        //worksheet.Cells[row + 31, col + i].Style.Font.Color.SetColor(Cores.corSecundáriaText);
                        //worksheet.Cells[row + 31, col + i].Style.Font.Italic = true;
                        //worksheet.Cells[row + 31, col + i].Style.Numberformat.Format = "0.00%";
                    }


                    if (i > 0 && i < numberOfYears / 2)
                    {
                        worksheet.Columns[col + i].OutlineLevel = 1;
                        worksheet.Columns[col + i].Collapsed = true;
                    }
                }
                else
                {
                    string columnLeft = columnLetter.GetExcelColumnName(col + i - 1);
                    string column = columnLetter.GetExcelColumnName(col + i);
                    string LastHistoricalColumn = columnLetter.GetExcelColumnName(numberOfYears + 2);

                    worksheet.Cells[row + 1, col + i].Formula = "="+ columnLeft + "3+1";
                    worksheet.Cells[row + 1, col + i].Style.Font.Bold = true;
                    worksheet.Cells[row + 1, col + i].Style.Border.Bottom.Style = ExcelBorderStyle.Thin;
                    worksheet.Columns[col + i].Width = 12;


                    if (i<numberOfYears+7)
                    {
                        string revenueOne = null;
                        string Cogs = null;
                        string OperatingCosts = null;
                        string CapexOne = null;
                        string intangibleOne = null;
                        string NWCOne = null;

                        if (i < numberOfYears + 2)
                        {
                            revenueOne = "'Assumptions'!C5";
                            Cogs = "'Assumptions'!C6";
                            OperatingCosts = "'Assumptions'!C7";
                            CapexOne = "'Assumptions'!C8";
                            intangibleOne = "'Assumptions'!C9";
                            NWCOne = "'Assumptions'!C10";


                        }
                        else
                        {
                            Cogs = "'Assumptions'!D6";
                            revenueOne = "'Assumptions'!D5";
                            OperatingCosts = "'Assumptions'!D7";
                            CapexOne = "'Assumptions'!D8";
                            intangibleOne = "'Assumptions'!D9";
                            NWCOne = "'Assumptions'!D10";
                        }
                        //Revenue
                        worksheet.Cells[row + 3, col + i].Formula = "=" + columnLeft + "5*(" + column + "6 +1)";
                        worksheet.Cells[row + 3, col + i].Style.Numberformat.Format = "#,##0;(#,##0);-";
                        worksheet.Cells[row + 3, col + i].Style.Font.Color.SetColor(Cores.CorTextoProjecoes);
                        

                        worksheet.Cells[row + 4, col + i].Formula = "=" + revenueOne;
                        worksheet.Cells[row + 4, col + i].Style.Font.Color.SetColor(Cores.corSecundáriaText);
                        worksheet.Cells[row + 4, col + i].Style.Font.Italic = true;
                        worksheet.Cells[row + 4, col + i].Style.Numberformat.Format = "0.00%";



                        //GCOGS
                        AuxProjectionValuesLessThanFive(worksheet, row + 3, col, i, Cogs);

                        //Operating Costs
                        AuxProjectionValuesLessThanFive(worksheet, row + 6, col, i, OperatingCosts);

                        //Core EBIT
                        int rowRevenue = row + 3;
                        int rowCostOfRevenue = row + 6;
                        int rowOpCosts = row + 9;
                        int rowEBIT = row + 12;

                        int rowTaxes = row + 15;
                        int rowEBIAT = row + 18;
                        int rowInvestedCapital = row + 21;
                        int rowDPA = row + 24;

                       

                        worksheet.Cells[row + 12, col + i].Formula = "=" + column + rowRevenue + "-" + column + rowOpCosts + "-" + column + rowCostOfRevenue+"-(" + column + "26-" + LastHistoricalColumn + "26)";
                        worksheet.Cells[row + 12, col + i].Style.Numberformat.Format = "#,##0;(#,##0);-";
                        worksheet.Cells[row + 12, col + i].Style.Font.Bold = true;
                        worksheet.Cells[row + 12, col + i].Style.Fill.PatternType = ExcelFillStyle.Solid;
                        worksheet.Cells[row + 12, col + i].Style.Fill.BackgroundColor.SetColor(Cores.CorCinzentoBackground);

                        //Less Taxes
                        worksheet.Cells[row + 15, col + i].Formula = "=" + column + rowEBIT + "* Assumptions!H6" ;
                        worksheet.Cells[row + 15, col + i].Style.Numberformat.Format = "#,##0;(#,##0);-";
                        worksheet.Cells[row + 16, col + i].Formula = "=" + column + rowTaxes + "/" + column + rowEBIT;
                        worksheet.Cells[row + 16, col + i].Style.Numberformat.Format = "0.00%";
                        worksheet.Cells[row + 16, col + i].Style.Font.Color.SetColor(Cores.corSecundáriaText);
                        worksheet.Cells[row + 16, col + i].Style.Font.Italic = true;

                        //Earnings before interest adjusted for taxes
                        worksheet.Cells[row + 18, col + i].Formula = "=" + column + rowEBIT + "-" + column + rowTaxes;
                        worksheet.Cells[row + 18, col + i].Style.Numberformat.Format = "#,##0;(#,##0);-";
                        worksheet.Cells[row + 18, col + i].Style.Fill.PatternType = ExcelFillStyle.Solid;
                        worksheet.Cells[row + 18, col + i].Style.Fill.BackgroundColor.SetColor(Cores.corSecundária);
                        worksheet.Cells[row + 18, col + i].Style.Font.Bold = true;

                        ////Capex
                        //worksheet.Cells[row + 21, col + i].Formula = "=" + columnLeft + "23*(" + column + "24 +1)";
                        //worksheet.Cells[row + 21, col + i].Style.Numberformat.Format = "#,##0;(#,##0);-";
                        //worksheet.Cells[row + 21, col + i].Style.Font.Color.SetColor(Cores.CorTextoProjecoes);


                        //worksheet.Cells[row + 22, col + i].Formula = "=" + CapexOne;
                        //worksheet.Cells[row + 22, col + i].Style.Font.Color.SetColor(Cores.corSecundáriaText);
                        //worksheet.Cells[row + 22, col + i].Style.Font.Italic = true;
                        //worksheet.Cells[row + 22, col + i].Style.Numberformat.Format = "0.00%";


                        ////Intangible
                        //worksheet.Cells[row + 24, col + i].Formula = "=" + columnLeft + "26*(" + column + "27 +1)";
                        //worksheet.Cells[row + 24, col + i].Style.Numberformat.Format = "#,##0;(#,##0);-";
                        //worksheet.Cells[row + 24, col + i].Style.Font.Color.SetColor(Cores.CorTextoProjecoes);


                        //worksheet.Cells[row + 25, col + i].Formula = "=" + intangibleOne;
                        //worksheet.Cells[row + 25, col + i].Style.Font.Color.SetColor(Cores.corSecundáriaText);
                        //worksheet.Cells[row + 25, col + i].Style.Font.Italic = true;
                        //worksheet.Cells[row + 25, col + i].Style.Numberformat.Format = "0.00%";


                        //Net Investment
                        worksheet.Cells[row + 21, col + i].Formula = "='Auxiliar'!" + column + "33";
                        worksheet.Cells[row + 21, col + i].Style.Numberformat.Format = "#,##0;(#,##0);-";
                        worksheet.Cells[row + 21, col + i].Style.Font.Color.SetColor(Cores.CorTextoProjecoes);


                        worksheet.Cells[row + 22, col + i].Formula = "=(" + column + "23/" + columnLeft + "23)-1";
                        worksheet.Cells[row + 22, col + i].Style.Font.Color.SetColor(Cores.corSecundáriaText);
                        worksheet.Cells[row + 22, col + i].Style.Font.Italic = true;
                        worksheet.Cells[row + 22, col + i].Style.Numberformat.Format = "0.00%";

                        //DP&A
                        worksheet.Cells[row + 24, col + i].Formula = "='Auxiliar'!" + column + "13";
                        worksheet.Cells[row + 24, col + i].Style.Numberformat.Format = "#,##0;(#,##0);-";
                        worksheet.Cells[row + 24, col + i].Style.Font.Color.SetColor(Cores.CorTextoProjecoes);


                        worksheet.Cells[row + 25, col + i].Formula = "='Assumptions'!H8";
                        worksheet.Cells[row + 25, col + i].Style.Font.Color.SetColor(Cores.corSecundáriaText);
                        worksheet.Cells[row + 25, col + i].Style.Font.Italic = true;
                        worksheet.Cells[row + 25, col + i].Style.Numberformat.Format = "0.00%";

                        ////NWC
                        //worksheet.Cells[row + 30, col + i].Formula = "=" + columnLeft + "32*(" + column + "33 +1)";
                        //worksheet.Cells[row + 30, col + i].Style.Numberformat.Format = "#,##0;(#,##0);-";
                        //worksheet.Cells[row + 30, col + i].Style.Font.Color.SetColor(Cores.CorTextoProjecoes);


                        //worksheet.Cells[row + 31, col + i].Formula = "=" + NWCOne;
                        //worksheet.Cells[row + 31, col + i].Style.Font.Color.SetColor(Cores.corSecundáriaText);
                        //worksheet.Cells[row + 31, col + i].Style.Font.Italic = true;
                        //worksheet.Cells[row + 31, col + i].Style.Numberformat.Format = "0.00%";


                        //FCF
                        worksheet.Cells[row + 27, col + i].Formula = "=" + column + rowEBIAT + "-" + column + rowInvestedCapital + "+" + column + rowDPA;
                        worksheet.Cells[row + 27, col + i].Style.Numberformat.Format = "#,##0;(#,##0);-";
                        worksheet.Cells[row + 27, col + i].Style.Fill.PatternType = ExcelFillStyle.Solid;
                        worksheet.Cells[row + 27, col + i].Style.Fill.BackgroundColor.SetColor(Cores.corSecundária);
                        worksheet.Cells[row + 27, col + i].Style.Font.Bold = true;

                    }
                    else
                    {
                        string revenueTwo = null;
                        if (i<numberOfYears+10)
                        {
                            revenueTwo = "'Assumptions'!E5";
                        }
                        else
                        {
                            revenueTwo = "'Assumptions'!M5";
                        }

                        //Revenue
                        worksheet.Cells[row + 3, col + i].Formula = "=" + columnLeft + "5*(" + column + "6 +1)";
                        worksheet.Cells[row + 3, col + i].Style.Numberformat.Format = "#,##0;(#,##0);-";
                        worksheet.Cells[row + 3, col + i].Style.Font.Color.SetColor(Cores.CorTextoProjecoes);
                        
                        worksheet.Cells[row + 4, col + i].Formula = "=" + revenueTwo;
                        worksheet.Cells[row + 4, col + i].Style.Font.Color.SetColor(Cores.corSecundáriaText);
                        worksheet.Cells[row + 4, col + i].Style.Font.Italic = true;
                        worksheet.Cells[row + 4, col + i].Style.Numberformat.Format = "0.00%";


                        //COGS
                        AuxProjectionValuesMoreThanFive(worksheet, row + 3, col, i, "'Assumptions'!E6");

                        //Operating costs
                        AuxProjectionValuesMoreThanFive(worksheet, row + 6, col, i, "'Assumptions'!E7");

                        //Core EBIT
                        int rowRevenue = row + 3;
                        int rowCostOfRevenue = row + 6;
                        int rowOpCosts = row + 9;
                        int rowEBIT = row + 12;
                        int rowTaxes = row + 15;
                        int rowEBIAT = row + 18;
                        int rowInvestedCapital = row + 21;
                        int rowDPA = row + 24;


                        string CapexTwo = "'Assumptions'!E8";
                        string intangibleTwo = "'Assumptions'!E9";
                        string NWCTwo = "'Assumptions'!E10";

                        worksheet.Cells[row + 12, col + i].Formula = "=" + column + rowRevenue + "-" + column + rowOpCosts + "-" + column + rowCostOfRevenue + "-(" + column + "26-" + LastHistoricalColumn + "26)";
                        worksheet.Cells[row + 12, col + i].Style.Numberformat.Format = "#,##0;(#,##0);-";
                        worksheet.Cells[row + 12, col + i].Style.Font.Bold = true;
                        worksheet.Cells[row + 12, col + i].Style.Fill.PatternType = ExcelFillStyle.Solid;
                        worksheet.Cells[row + 12, col + i].Style.Fill.BackgroundColor.SetColor(Cores.CorCinzentoBackground);

                        //Less Taxes
                        worksheet.Cells[row + 15, col + i].Formula = "=" + column + rowEBIT + "*Assumptions!H6";
                        worksheet.Cells[row + 15, col + i].Style.Numberformat.Format = "#,##0;(#,##0);-";
                        worksheet.Cells[row + 16, col + i].Formula = "=" + column + rowTaxes + "/" + column + rowEBIT;
                        worksheet.Cells[row + 16, col + i].Style.Numberformat.Format = "0.00%";
                        worksheet.Cells[row + 16, col + i].Style.Font.Color.SetColor(Cores.corSecundáriaText);
                        worksheet.Cells[row + 16, col + i].Style.Font.Italic = true;


                        //Earnings before interest adjusted for taxes
                        worksheet.Cells[row + 18, col + i].Formula = "=" + column + rowEBIT + "-" + column + rowTaxes;
                        worksheet.Cells[row + 18, col + i].Style.Numberformat.Format = "#,##0;(#,##0);-";
                        worksheet.Cells[row + 18, col + i].Style.Fill.PatternType = ExcelFillStyle.Solid;
                        worksheet.Cells[row + 18, col + i].Style.Fill.BackgroundColor.SetColor(Cores.corSecundária);
                        worksheet.Cells[row + 18, col + i].Style.Font.Bold = true;

                        ////Capex
                        //worksheet.Cells[row + 21, col + i].Formula = "=" + columnLeft + "23*(" + column + "24 +1)";
                        //worksheet.Cells[row + 21, col + i].Style.Numberformat.Format = "#,##0;(#,##0);-";
                        //worksheet.Cells[row + 21, col + i].Style.Font.Color.SetColor(Cores.CorTextoProjecoes);


                        //worksheet.Cells[row + 22, col + i].Formula = "=" + CapexTwo;
                        //worksheet.Cells[row + 22, col + i].Style.Font.Color.SetColor(Cores.corSecundáriaText);
                        //worksheet.Cells[row + 22, col + i].Style.Font.Italic = true;
                        //worksheet.Cells[row + 22, col + i].Style.Numberformat.Format = "0.00%";


                        ////Intangible
                        //worksheet.Cells[row + 24, col + i].Formula = "=" + columnLeft + "26*(" + column + "27 +1)";
                        //worksheet.Cells[row + 24, col + i].Style.Numberformat.Format = "#,##0;(#,##0);-";
                        //worksheet.Cells[row + 24, col + i].Style.Font.Color.SetColor(Cores.CorTextoProjecoes);


                        //worksheet.Cells[row + 25, col + i].Formula = "=" + intangibleTwo;
                        //worksheet.Cells[row + 25, col + i].Style.Font.Color.SetColor(Cores.corSecundáriaText);
                        //worksheet.Cells[row + 25, col + i].Style.Font.Italic = true;
                        //worksheet.Cells[row + 25, col + i].Style.Numberformat.Format = "0.00%";

                        //Net Investment
                        worksheet.Cells[row + 21, col + i].Formula = "='Auxiliar'!" + column + "33";
                        worksheet.Cells[row + 21, col + i].Style.Numberformat.Format = "#,##0;(#,##0);-";
                        worksheet.Cells[row + 21, col + i].Style.Font.Color.SetColor(Cores.CorTextoProjecoes);


                        worksheet.Cells[row + 22, col + i].Formula = "=(" + column + "23/" + columnLeft + "23)-1";
                        worksheet.Cells[row + 22, col + i].Style.Font.Color.SetColor(Cores.corSecundáriaText);
                        worksheet.Cells[row + 22, col + i].Style.Font.Italic = true;
                        worksheet.Cells[row + 22, col + i].Style.Numberformat.Format = "0.00%";

                        //DP&A
                        worksheet.Cells[row + 24, col + i].Formula = "='Auxiliar'!" + column + "13";
                        worksheet.Cells[row + 24, col + i].Style.Numberformat.Format = "#,##0;(#,##0);-";
                        worksheet.Cells[row + 24, col + i].Style.Font.Color.SetColor(Cores.CorTextoProjecoes);


                        worksheet.Cells[row + 25, col + i].Formula = "='Assumptions'!H8";
                        worksheet.Cells[row + 25, col + i].Style.Font.Color.SetColor(Cores.corSecundáriaText);
                        worksheet.Cells[row + 25, col + i].Style.Font.Italic = true;
                        worksheet.Cells[row + 25, col + i].Style.Numberformat.Format = "0.00%";


                        ////NWC
                        //worksheet.Cells[row + 30, col + i].Formula = "=" + columnLeft + "32*(" + column + "33 +1)";
                        //worksheet.Cells[row + 30, col + i].Style.Numberformat.Format = "#,##0;(#,##0);-";
                        //worksheet.Cells[row + 30, col + i].Style.Font.Color.SetColor(Cores.CorTextoProjecoes);


                        //worksheet.Cells[row + 31, col + i].Formula = "=" + NWCTwo;
                        //worksheet.Cells[row + 31, col + i].Style.Font.Color.SetColor(Cores.corSecundáriaText);
                        //worksheet.Cells[row + 31, col + i].Style.Font.Italic = true;
                        //worksheet.Cells[row + 31, col + i].Style.Numberformat.Format = "0.00%";

                        //FCF
                        worksheet.Cells[row + 27, col + i].Formula = "=" + column + rowEBIAT + "-" + column + rowInvestedCapital + "+" + column + rowDPA;
                        worksheet.Cells[row + 27, col + i].Style.Numberformat.Format = "#,##0;(#,##0);-";
                        worksheet.Cells[row + 27, col + i].Style.Fill.PatternType = ExcelFillStyle.Solid;
                        worksheet.Cells[row + 27, col + i].Style.Fill.BackgroundColor.SetColor(Cores.corSecundária);
                        worksheet.Cells[row + 27, col + i].Style.Font.Bold = true;
                    }

                }
            }


        }
        public void AuxHistoricalDescription(ExcelWorksheet worksheet, int row, int col, int i, int numberCaption)
        {
            worksheet.Cells[row + 3, col + i].Formula = "='P&L'!B" + numberCaption;
            
            worksheet.Cells[row + 4, col + i].Value = "% revenues";
            worksheet.Cells[row + 4, col + i].Style.Font.Color.SetColor(Cores.corSecundáriaText);
            worksheet.Cells[row + 4, col + i].Style.Font.Italic = true;
        }

        public void AuxHistoricalValues(ExcelWorksheet worksheet, int row, int col, int i, int numberCaption, string Nature)
        {
            ExcelNextCol columnName = new ExcelNextCol();
            string column = columnName.GetExcelColumnName(col+i);
            string columnLeft = columnName.GetExcelColumnName(col+i-1);
            int rowCell = row + 3;

            if (Nature == "N")
            {
                worksheet.Cells[row + 3, col + i].Formula = "=-'P&L'!" + column + numberCaption;
            }
            else
            {
                worksheet.Cells[row + 3, col + i].Formula = "='P&L'!" + column + numberCaption;
            }
            
            worksheet.Cells[row + 3, col + i].Style.Numberformat.Format = "#,##0;(#,##0);-";
            worksheet.Cells[row + 4, col + i].Formula = "=" + column + rowCell + "/" + column + 5;
            worksheet.Cells[row + 4, col + i].Style.Numberformat.Format = "0.00%";
            worksheet.Cells[row + 4, col + i].Style.Font.Color.SetColor(Cores.corSecundáriaText);
            worksheet.Cells[row + 4, col + i].Style.Font.Italic = true;
        }

        public void AuxProjectionValuesLessThanFive(ExcelWorksheet worksheet, int row, int col, int i, string RevenueOne)
        {
            ExcelNextCol columnName = new ExcelNextCol();
            string column = columnName.GetExcelColumnName(col + i);
            string columnLeft = columnName.GetExcelColumnName(col + i - 1);
            int rowCell = row + 3;
            int rowCellPlusOne = row + 4;


            worksheet.Cells[row + 3, col + i].Formula = "=" + column + 5 + "*" + column + rowCellPlusOne;
            worksheet.Cells[row + 3, col + i].Style.Numberformat.Format = "#,##0;(#,##0);-";
            worksheet.Cells[row + 3, col + i].Style.Font.Color.SetColor(Cores.CorTextoProjecoes);
            
            worksheet.Cells[row + 4, col + i].Formula = "=" + columnLeft + rowCellPlusOne + "*(1+" + RevenueOne+ ")";
            worksheet.Cells[row + 4, col + i].Style.Numberformat.Format = "0.00%";
            worksheet.Cells[row + 4, col + i].Style.Font.Color.SetColor(Cores.corSecundáriaText);
            worksheet.Cells[row + 4, col + i].Style.Font.Italic = true;
        }
        public void AuxProjectionValuesMoreThanFive(ExcelWorksheet worksheet, int row, int col, int i, string RevenueTwo)
        {
            ExcelNextCol columnName = new ExcelNextCol();
            string column = columnName.GetExcelColumnName(col + i);
            string columnLeft = columnName.GetExcelColumnName(col + i - 1);
            int rowCell = row + 3;
            int rowCellPlusOne = row + 4;


            worksheet.Cells[row + 3, col + i].Formula = "=" + column + 5 + "*" + column + rowCellPlusOne;
            worksheet.Cells[row + 3, col + i].Style.Numberformat.Format = "#,##0;(#,##0);-";
            worksheet.Cells[row + 3, col + i].Style.Font.Color.SetColor(Cores.CorTextoProjecoes);
            worksheet.Cells[row + 4, col + i].Formula = "=" + columnLeft + rowCellPlusOne + "*(1+" + RevenueTwo + ")";
            worksheet.Cells[row + 4, col + i].Style.Numberformat.Format = "0.00%";
            worksheet.Cells[row + 4, col + i].Style.Font.Color.SetColor(Cores.corSecundáriaText);
            worksheet.Cells[row + 4, col + i].Style.Font.Italic = true;
        }
        public double GetTaxByYear(Taxes tax, int LastYear, int numberOfYears, int i)
        {
            int year = LastYear - numberOfYears + i - 1;

            string aux = "year" + year;

            try
            {
                string value = GetPropertyValue(tax, aux).ToString();

                NumberFormatInfo provider = new NumberFormatInfo();
                provider.NumberDecimalSeparator = ".";

                double output = double.Parse(value, provider);

                return output;
            }
            catch (Exception)
            {

                return 0;
            }


        }

        public object GetPropertyValue(object obj, string propertyName)
        {
            Type type = obj.GetType();
            PropertyInfo property = type.GetProperty(propertyName);
            return property.GetValue(obj, null);
        }



    }
    public class Taxes
    {
        public string iso_2 { get; set; }
        public string iso_3 { get; set; }
        public string continent { get; set; }
        public string country { get; set; }
        public string year2008 { get; set; }
        public string year2009 { get; set; }
        public string year2010 { get; set; }
        public string year2011 { get; set; }
        public string year2012 { get; set; }
        public string year2013 { get; set; }
        public string year2014 { get; set; }
        public string year2015 { get; set; }
        public string year2016 { get; set; }
        public string year2017 { get; set; }
        public string year2018 { get; set; }
        public string year2019 { get; set; }
        public string year2020 { get; set; }
        public string year2021 { get; set; }
        public string year2022 { get; set; }


    }

}
