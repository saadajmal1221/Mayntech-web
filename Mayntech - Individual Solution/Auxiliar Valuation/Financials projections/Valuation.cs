using Mayntech___Individual_Solution.Auxiliar.Auxiliar.FMP;
using Mayntech___Individual_Solution.Auxiliar.Auxiliar.Other;
using Mayntech___Individual_Solution.Auxiliar_Analysis.Auxiliar.FMP;
using OfficeOpenXml;
using OfficeOpenXml.Style;
using System.Drawing;
using System.Globalization;

namespace Mayntech___Individual_Solution.Auxiliar_Valuation.Financials_projections
{
    public class Valuation
    {
        public void ValuationBuilder(ExcelPackage package, int numberOfYears, CompanyProfile companyOutlook, string currency, ForexList forex)
        {
            ExcelNextCol columnName = new ExcelNextCol();
            var worksheet = package.Workbook.Worksheets.Add("DCF Valuation");
            worksheet.View.ShowGridLines = false;
            worksheet.View.ZoomScale = 80;
            worksheet.View.FreezePanes(4, 3);
            int row = 2;
            int col = 2;

            string aux =  "USD/" + currency;

            Fx? FxRate = forex.forexList.FirstOrDefault(t => t.ticker == aux);

            if (FxRate == null)
            {

                FxRate = forex.forexList.FirstOrDefault(t => t.ticker == currency + "/USD");
            }

            double rate = 0;
            if (FxRate != null)
            {
                NumberFormatInfo provider = new NumberFormatInfo();
                provider.NumberDecimalSeparator = ".";

                double output1 = double.Parse(FxRate.bid, provider);
                double output2 = double.Parse(FxRate.ask, provider);

                rate = (output1 + output2) / 2;
            }
            

            for (int i = 0; i < 12; i++)
            {

                worksheet.Cells[row, col + i].Style.Fill.PatternType = ExcelFillStyle.Solid;
                worksheet.Cells[row, col + i].Style.Fill.BackgroundColor.SetColor(Cores.CorPrincipal);


                if (i == 0)
                {
                    worksheet.Cells[row, col].Value = "DCF";
                    worksheet.Cells[row, col].Style.Font.Bold = true;
                    worksheet.Cells[row, col].Style.Font.Color.SetColor(Color.White);

                    worksheet.Cells[row + 1, col + i].Formula = "='FCF Projections'!B3";
                    worksheet.Cells[row + 1, col + i].Style.Font.Color.SetColor(Cores.CorTexto);
                    worksheet.Cells[row + 1, col + i].Style.Border.Bottom.Style = ExcelBorderStyle.Thin;
                    worksheet.Cells[row + 1, col + i].Style.Font.Bold = true;

                    worksheet.Cells[row + 3, col + i].Value = "PV Future Cash flows";
                    worksheet.Cells[row + 4, col + i].Value = "Total PV";
                    worksheet.Cells[row + 5, col + i].Value = "Terminal Value";
                    worksheet.Cells[row + 6, col + i].Value = "PV of Terminal Value";
                    worksheet.Cells[row + 7, col + i].Value = "Enterprise Value";
                    worksheet.Cells[row + 8, col + i].Value = "Net Debt";
                    worksheet.Cells[row + 9, col + i].Value = "Minority Interests";
                    worksheet.Cells[row + 10, col + i].Value = "Equity Value";

                    worksheet.Cells[row + 12, col + i].Value = "Number of shares outstanding (in thousands)";
                    worksheet.Cells[row + 13, col + i].Value = "Market Price per share (in " + currency + ")";
                    worksheet.Cells[row + 13, col + i].Style.Font.Bold = true;

                    worksheet.Cells[row + 14, col + i].Value = "Price per share DCF (in " + currency + ")";
                    worksheet.Cells[row + 14, col + i].Style.Font.Bold = true;

                    if (FxRate != null)
                    {

                        worksheet.Cells[row + 16, col + i].Value = "Market price per share (in USD)";

                        worksheet.Cells[row + 17, col + i].Value = "Price per share DCF (in USD)";

                        worksheet.Cells[row + 19, col + i].Value = "Exchange Rate";

                    }

                    worksheet.Columns[col + i].Width = 45;
                }
                else
                {
                    string column = columnName.GetExcelColumnName(col + numberOfYears + i);
                    string columnShares = columnName.GetExcelColumnName(col + numberOfYears);
                    string Lastcolumn = columnName.GetExcelColumnName(col + 11);
                    string LastcolumnReported = columnName.GetExcelColumnName(col + 10);
                    string columnThis = columnName.GetExcelColumnName(col + i);
                    int auxPower = i - 1;

                    worksheet.Cells[row + 1, col + i].Formula = "='FCF Projections'!" + column + "3";
                    worksheet.Cells[row + 1, col + i].Style.Font.Color.SetColor(Cores.CorTexto);
                    worksheet.Cells[row + 1, col + i].Style.Border.Bottom.Style = ExcelBorderStyle.Thin;
                    worksheet.Cells[row + 1, col + i].Style.Font.Bold = true;

                    worksheet.Rows[row + 2].Height = 3;

                    //PV Future cash flows
                    worksheet.Cells[row + 3, col + i].Formula = "='FCF Projections'!" + column + "29/((1+'Assumptions'!H5)^" + auxPower + ")";
                    worksheet.Cells[row + 3, col + i].Style.Numberformat.Format = "#,##0;(#,##0);-";

                    if (i == 1)
                    {
                        //Total PV
                        worksheet.Cells[row + 4, col + i].Formula = "=SUM(C5:" + Lastcolumn +"5)" ;
                        worksheet.Cells[row + 4, col + i].Style.Numberformat.Format = "#,##0;(#,##0);-";

                        //terminal Value
                        worksheet.Cells[row + 5, col + i].Formula = "=(" + Lastcolumn + "5*(1+'Assumptions'!H16))/('Assumptions'!H5 - 'Assumptions'!M7)";
                        worksheet.Cells[row + 5, col + i].Style.Numberformat.Format = "#,##0;(#,##0);-";

                        //Present Value terminal Value
                        worksheet.Cells[row + 6, col + i].Formula = "=" + columnThis + "7/((1+'Assumptions'!H5)^10)";
                        worksheet.Cells[row + 6, col + i].Style.Numberformat.Format = "#,##0;(#,##0);-";

                        //enterprise Value
                        worksheet.Cells[row + 7, col + i].Formula = "=" + columnThis + "8 + " + columnThis + "6";
                        worksheet.Cells[row + 7, col + i].Style.Numberformat.Format = "#,##0;(#,##0);-";

                        //Net Debt
                        worksheet.Cells[row + 8, col + i].Formula = "='BS'!" + LastcolumnReported + "28 + 'BS'!" + LastcolumnReported + "35 - 'BS'!" + LastcolumnReported + "7";
                        worksheet.Cells[row + 8, col + i].Style.Numberformat.Format = "#,##0;(#,##0);-";

                        //Minority Interests
                        worksheet.Cells[row + 9, col + i].Formula = "='BS'!" + LastcolumnReported + "52";
                        worksheet.Cells[row + 9, col + i].Style.Numberformat.Format = "#,##0;(#,##0);-";

                        //Equity Value
                        worksheet.Cells[row + 10, col + i].Formula = "=" + columnThis + "9-" + columnThis + "10-" + columnThis + "11";
                        worksheet.Cells[row + 10, col + i].Style.Numberformat.Format = "#,##0;(#,##0);-";

                        //Number of shares outstanding
                        worksheet.Cells[row + 12, col + i].Formula = "='P&L'!" + columnShares + "39/1000";
                        worksheet.Cells[row + 12, col + i].Style.Numberformat.Format = "#,##0;(#,##0);-";

                        //Market Price per share
                        worksheet.Cells[row + 13, col + i].Value = companyOutlook.profile.price;
                        worksheet.Cells[row + 13, col + i].Style.Numberformat.Format = "#,##0.00;(#,##0.00);-";
                        worksheet.Cells[row + 13, col + i].Style.Font.Bold = true;


                        //DCF Price
                        worksheet.Cells[row + 14, col + i].Formula = "=MAX(" + columnThis + "12/" + columnThis + "14,0)";
                        worksheet.Cells[row + 14, col + i].Style.Numberformat.Format = "#,##0.00;(#,##0.00);-";
                        worksheet.Cells[row + 14, col + i].Style.Font.Bold = true;

                        if (FxRate != null)
                        {

                            //Market Price USD
                            worksheet.Cells[row + 16, col + i].Formula = "=" + columnThis + "15/" + columnThis + "21";
                            worksheet.Cells[row + 16, col + i].Style.Numberformat.Format = "#,##0.00;(#,##0.00);-";


                            //DCF Price dollars
                            worksheet.Cells[row + 17, col + i].Formula = "=" + columnThis + "16/" + columnThis + "21";
                            worksheet.Cells[row + 17, col + i].Style.Numberformat.Format = "#,##0.00;(#,##0.00);-";

                            //Exchange Rate
                            worksheet.Cells[row + 19, col + i].Value =  rate;
                            worksheet.Cells[row + 19, col + i].Style.Numberformat.Format = "#,##0.00;(#,##0.00);-";

                        }

                        worksheet.Columns[col + i].Width = 15;
                    }
                    if (i!=1)
                    {
                        worksheet.Columns[col + i].Width = 12;
                    }
                    
                }
            }
        }
    }
}
