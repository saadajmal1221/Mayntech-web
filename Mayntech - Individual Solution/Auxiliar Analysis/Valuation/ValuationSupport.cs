using Mayntech___Individual_Solution.Auxiliar.Analysis;
using Mayntech___Individual_Solution.Auxiliar.Auxiliar.FMP;
using Mayntech___Individual_Solution.Auxiliar.Auxiliar.Other;
using Mayntech___Individual_Solution.Auxiliar_Valuation.Financials_projections;
using Mayntech___Individual_Solution.Pages;
using OfficeOpenXml;
using OfficeOpenXml.Style;
using System.Drawing;
using System.Runtime.Intrinsics.X86;
using System.Xml.Linq;

namespace Mayntech___Individual_Solution.Auxiliar_Analysis.Valuation
{
    public class ValuationSupport
    {
        public void ValuationSupportConstruction(ExcelPackage package, int numberOfYears, Taxes tax, string calendarYear, 
            List<FinancialStatements> balanceSheet)
        {
            int LastYear = int.Parse(calendarYear);
            ExcelNextCol columnLetter = new ExcelNextCol();
            var worksheet = package.Workbook.Worksheets.Add("Valuation Support");
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

            for (int i = 0; i < numberOfYears + 1; i++)
            {
                worksheet.Cells[row, col + i].Style.Fill.PatternType = ExcelFillStyle.Solid;
                worksheet.Cells[row, col + i].Style.Fill.BackgroundColor.SetColor(Cores.CorPrincipal);
                FCFProjections fcf = new FCFProjections();

                if (i < 1)
                {
                    worksheet.Cells[row, col + i].Value = "Free cash flow";
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
                    
                    fcf.AuxHistoricalDescription(worksheet, row + 3, col, i, 6);

                    //OperatingCosts
                    fcf.AuxHistoricalDescription(worksheet, row + 6, col, i, 10);

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




                    //FCF
                    worksheet.Cells[row + 27, col + i].Value = "Free Cash Flow";
                    worksheet.Cells[row + 27, col + i].Style.Fill.PatternType = ExcelFillStyle.Solid;
                    worksheet.Cells[row + 27, col + i].Style.Fill.BackgroundColor.SetColor(Cores.corSecundária);
                    worksheet.Cells[row + 27, col + i].Style.Font.Bold = true;




                    //Invested Capital
                    worksheet.Cells[row + 30, col + i].Value = "Invested Capital";
                    worksheet.Cells[row + 30, col + i].Style.Font.Color.SetColor(Cores.CorTexto);
                    worksheet.Cells[row + 30, col + i].Style.Border.Bottom.Style = ExcelBorderStyle.Thin;
                    worksheet.Cells[row + 30, col + i].Style.Font.Bold = true;


                    worksheet.Cells[row + 31, col + i].Value = "PP&E";
                    worksheet.Cells[row + 32, col + i].Value = "% Change";
                    worksheet.Cells[row + 32, col + i].Style.Font.Color.SetColor(Cores.corSecundáriaText);

                    worksheet.Cells[row + 34, col + i].Value = "Intangibles";
                    worksheet.Cells[row + 35, col + i].Value = "% Change";
                    worksheet.Cells[row + 35, col + i].Style.Font.Color.SetColor(Cores.corSecundáriaText);

                    worksheet.Cells[row + 37, col + i].Value = "Goodwill";
                    worksheet.Cells[row + 38, col + i].Value = "% Change";
                    worksheet.Cells[row + 38, col + i].Style.Font.Color.SetColor(Cores.corSecundáriaText);

                    worksheet.Cells[row + 40, col + i].Value = "NWC";
                    worksheet.Cells[row + 41, col + i].Value = "% revenues";
                    worksheet.Cells[row + 41, col + i].Style.Font.Color.SetColor(Cores.corSecundáriaText);

                    worksheet.Cells[row + 43, col + i].Value = "Invested Capital";
                    worksheet.Cells[row + 43, col + i].Style.Fill.PatternType = ExcelFillStyle.Solid;
                    worksheet.Cells[row + 43, col + i].Style.Fill.BackgroundColor.SetColor(Cores.corSecundária);
                    worksheet.Cells[row + 43, col + i].Style.Font.Bold = true;

                    worksheet.Cells[row + 45, col + i].Value = "Net Investment";
                    worksheet.Cells[row + 45, col + i].Style.Fill.PatternType = ExcelFillStyle.Solid;
                    worksheet.Cells[row + 45, col + i].Style.Fill.BackgroundColor.SetColor(Cores.CorCinzentoBackground);
                    worksheet.Cells[row + 45, col + i].Style.Font.Bold = true;


                    worksheet.Columns[col + i].Width = 25;

                }
                else
                {


                    string columnLeft = columnLetter.GetExcelColumnName(col + i - 1);
                    string column = columnLetter.GetExcelColumnName(col + i);
                    worksheet.Cells[row + 1, col + i].Formula = "=YEAR('P&L'!" + column + "3)";
                    worksheet.Cells[row + 1, col + i].Style.Font.Bold = true;
                    worksheet.Cells[row + 1, col + i].Style.Border.Bottom.Style = ExcelBorderStyle.Thin;
                    worksheet.Columns[col + i].Width = 12;

                    //Revenue
                    worksheet.Cells[row + 3, col + i].Formula = "='P&L'!" + column + "5";
                    worksheet.Cells[row + 3, col + i].Style.Numberformat.Format = "#,##0;(#,##0);-";

                    List<FinancialStatements> incomeStatement = new List<FinancialStatements>();


                    //Comentário Revenue


                    //Gross Profit
                    fcf.AuxHistoricalValues(worksheet, row + 3, col, i, 6, "N");

                    //Operating Costs
                    fcf.AuxHistoricalValues(worksheet, row + 6, col, i, 10, "N");

                    //Core EBIT
                    int rowRevenue = row + 3;
                    int rowCostOfRevenue = row + 6;
                    int rowOpCosts = row + 9;
                    int rowEBIT = row + 12;
                    int rowTaxes = row + 15;
                    int rowTaxeRate = row + 16;
                    int rowEBIAT = row + 18;
                    int rowInvestedCapital = row + 21;
                    int rowDPA = row + 24;
                    FCFProjections fCFProjections = new FCFProjections();
                    

                    worksheet.Cells[row + 12, col + i].Formula = "=" + column + rowRevenue + "-" + column + rowOpCosts + "-" + column + rowCostOfRevenue;
                    worksheet.Cells[row + 12, col + i].Style.Numberformat.Format = "#,##0;(#,##0);-";
                    worksheet.Cells[row + 12, col + i].Style.Font.Bold = true;
                    worksheet.Cells[row + 12, col + i].Style.Fill.PatternType = ExcelFillStyle.Solid;
                    worksheet.Cells[row + 12, col + i].Style.Fill.BackgroundColor.SetColor(Cores.CorCinzentoBackground);

                    //Less Taxes
                    worksheet.Cells[row + 15, col + i].Formula = "=" + column + rowEBIT + "*" + column + rowTaxeRate;
                    worksheet.Cells[row + 15, col + i].Style.Numberformat.Format = "#,##0;(#,##0);-";
                    worksheet.Cells[row + 16, col + i].Value = fCFProjections.GetTaxByYear(tax, LastYear, numberOfYears, i) / 100;
                    worksheet.Cells[row + 16, col + i].Style.Numberformat.Format = "0.00%";
                    worksheet.Cells[row + 16, col + i].Style.Font.Color.SetColor(Cores.corSecundáriaText);
                    worksheet.Cells[row + 16, col + i].Style.Font.Italic = true;

                    //Earnings before interest adjusted for taxes
                    worksheet.Cells[row + 18, col + i].Formula = "=" + column + rowEBIT + "-" + column + rowTaxes;
                    worksheet.Cells[row + 18, col + i].Style.Numberformat.Format = "#,##0;(#,##0);-";
                    worksheet.Cells[row + 18, col + i].Style.Fill.PatternType = ExcelFillStyle.Solid;
                    worksheet.Cells[row + 18, col + i].Style.Fill.BackgroundColor.SetColor(Cores.corSecundária);
                    worksheet.Cells[row + 18, col + i].Style.Font.Bold = true;


                    //Net Investment
                    worksheet.Cells[row + 21, col + i].Formula = "=" + column + "47";
                    worksheet.Cells[row + 21, col + i].Style.Numberformat.Format = "#,##0;(#,##0);-";

                    worksheet.Cells[row + 22, col + i].Formula = "=(" + column + "23/" + columnLeft + "23)-1";
                    worksheet.Cells[row + 22, col + i].Style.Font.Color.SetColor(Cores.corSecundáriaText);
                    worksheet.Cells[row + 22, col + i].Style.Font.Italic = true;
                    worksheet.Cells[row + 22, col + i].Style.Numberformat.Format = "0.00%";

                    //DP&A
                    worksheet.Cells[row + 24, col + i].Formula = "=-'P&L'!" + column + "32";
                    worksheet.Cells[row + 24, col + i].Style.Numberformat.Format = "#,##0;(#,##0);-";

                    worksheet.Cells[row + 25, col + i].Formula = "=-'P&L'!" + column + "32/('BS'!" + column + "15+'BS'!" + column + "17)";
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

                    if (i > 1)
                    {

                        worksheet.Cells[row + 4, col + i].Formula = "=(" + column + "5/" + columnLeft + "5)-1";
                        worksheet.Cells[row + 4, col + i].Style.Font.Color.SetColor(Cores.corSecundáriaText);
                        worksheet.Cells[row + 4, col + i].Style.Font.Italic = true;
                        worksheet.Cells[row + 4, col + i].Style.Numberformat.Format = "0.00%";

                    }

                    //Invested Capital
                    worksheet.Cells[row + 30, col + i].Style.Border.Bottom.Style = ExcelBorderStyle.Thin;
                    worksheet.Cells[row + 30, col + i].Style.Font.Bold = true;

                    //PP&E

                    worksheet.Cells[row + 31, col + i].Formula = "='BS'!" + column + "15";
                    worksheet.Cells[row + 31, col + i].Style.Numberformat.Format = "#,##0;(#,##0);-";

                    if (i > 1)
                    {
                        worksheet.Cells[row + 32, col + i].Formula = "=(" + column + "33/" + columnLeft + "33)-1";
                        worksheet.Cells[row + 32, col + i].Style.Numberformat.Format = "0.00%";
                        worksheet.Cells[row + 32, col + i].Style.Font.Color.SetColor(Cores.corSecundáriaText);
                    }


                    //Intangibles

                    worksheet.Cells[row + 34, col + i].Formula = "='BS'!" + column + "17";
                    worksheet.Cells[row + 34, col + i].Style.Numberformat.Format = "#,##0;(#,##0);-";

                    if (i > 1)
                    {
                        worksheet.Cells[row + 35, col + i].Formula = "=(" + column + "36/" + columnLeft + "36)-1";
                        worksheet.Cells[row + 35, col + i].Style.Numberformat.Format = "0.00%";
                        worksheet.Cells[row + 35, col + i].Style.Font.Color.SetColor(Cores.corSecundáriaText);
                    }


                    //Goodwill

                    worksheet.Cells[row + 37, col + i].Formula = "='BS'!" + column + "16";
                    worksheet.Cells[row + 37, col + i].Style.Numberformat.Format = "#,##0;(#,##0);-";

                    if (i > 1)
                    {
                        worksheet.Cells[row + 38, col + i].Formula = "=(" + column + "39/" + columnLeft + "39) - 1";
                        worksheet.Cells[row + 38, col + i].Style.Numberformat.Format = "0.00%";
                        worksheet.Cells[row + 38, col + i].Style.Font.Color.SetColor(Cores.corSecundáriaText);
                    }




                    //NWC
                    if (investedCapital.Min() < 0)
                    {
                        worksheet.Cells[row + 40, col + i].Formula = "=('BS'!" + column + "9 + 'BS'!" + column + "10)-('BS'!" + column + "27 + 'BS'!" + column + "30)";
                    }
                    else
                    {
                        worksheet.Cells[row + 40, col + i].Formula = "=('BS'!" + column + "9 + 'BS'!" + column + "10 +  'BS'!" + column + "11)-('BS'!" + column + "27 + 'BS'!" + column + "30+ 'BS'!" + column + "31)";
                    }

                    worksheet.Cells[row + 40, col + i].Style.Numberformat.Format = "#,##0;(#,##0);-";

                    worksheet.Cells[row + 41, col + i].Formula = "=" + column + "42/'P&L'!" + column + "5";
                    worksheet.Cells[row + 41, col + i].Style.Numberformat.Format = "0.00%";
                    worksheet.Cells[row + 41, col + i].Style.Font.Color.SetColor(Cores.corSecundáriaText);

                    worksheet.Columns[col + i].Width = 12;


                    //Operational Invested capital
                    worksheet.Cells[row + 43, col + i].Formula = "=" + column + "33 +" + column + "36 +" + column + "42";
                    worksheet.Cells[row + 43, col + i].Style.Numberformat.Format = "#,##0;(#,##0);-";
                    worksheet.Cells[row + 43, col + i].Style.Fill.PatternType = ExcelFillStyle.Solid;
                    worksheet.Cells[row + 43, col + i].Style.Fill.BackgroundColor.SetColor(Cores.corSecundária);
                    worksheet.Cells[row + 43, col + i].Style.Font.Bold = true;

                    worksheet.Cells[row + 45, col + i].Style.Fill.PatternType = ExcelFillStyle.Solid;
                    worksheet.Cells[row + 45, col + i].Style.Fill.BackgroundColor.SetColor(Cores.CorCinzentoBackground);

                    if (i > 1)
                    {
                        worksheet.Cells[row + 45, col + i].Formula = "=" + column + "45-" + columnLeft + "45";
                        worksheet.Cells[row + 45, col + i].Style.Numberformat.Format = "#,##0;(#,##0);-";
                        worksheet.Cells[row + 45, col + i].Style.Font.Bold = true;

                    }

                }

            }
        }

        public List<string> Analyse(List<double> inputList)
        {
            List<double> xvalues = new List<double>();

            List<string> result = new List<string>();

            for (int i = 0; i < inputList.Count(); i++)
            {
                xvalues.Add(i);
            }
            LinearRegression linearRegression = new LinearRegression();
            List<double> OutputLR = linearRegression.LinearRegressionCalculation(xvalues, inputList);

            //Slope, Intercept, R-Squared

            double cagr = Cagr(inputList);
            double coefficient = CoefficientVariation(inputList);


            if (OutputLR[2] > 0.4)
            {
                if (OutputLR[0]>0 && cagr>0)
                {
                    result.Add("increase");
                }
                else if (OutputLR[0] > 0 && cagr < 0 || OutputLR[0] == 0)
                {
                    result.Add("unclear");
                }
                else if (OutputLR[0] < 0 && cagr<0)
                {
                    result.Add("decrease");
                }
                else
                {
                    result.Add("unclear");
                }
            }
            else
            {
                if (OutputLR[0] > 0 && cagr > 0.05)
                {
                    result.Add("increase");
                }
                else if (OutputLR[0] < 0 && cagr < -0.05)
                {
                    result.Add("decrease");
                }
                else
                {
                    result.Add("unclear");
                }
            }


            if (coefficient > 0.3)
            {
                result.Add("high volatility");
            }
            else
            {
                result.Add("low volatility");
            }

            return result;
        }
        public double Cagr(List<double> inputList)
        {
            double firstValue = inputList[0];
            double LastValue = inputList[inputList.Count - 1];

            double output = Math.Pow(LastValue/firstValue, (double)1/(double)(inputList.Count - 1))-1;

            if (firstValue<0 && LastValue<0)
            {

                output = -Math.Pow(LastValue / firstValue, (double)1 / (double)(inputList.Count - 1))+1;

            }
            else if (firstValue < 0 && LastValue > 0)
            {
                output = 0.06;
            }
            else if (firstValue > 0 && LastValue < 0)
            {
                output = -0.06;
            }

            return output;
        }

        public double CoefficientVariation(List<double> inputList)
        {

            double average = inputList.Average();
            double sumOfSquaresOfDifferences = inputList.Select(val => (val - average) * (val - average)).Sum();
            double sd = Math.Sqrt(sumOfSquaresOfDifferences / inputList.Count());

            double result = sd / average;

            return result;
        }

        public string commentsRevenue(List<string> revenue)
        {
            List<string> aux = new List<string>(); ;

            if (revenue[0] == "increase")
            {
                aux.Add("The company has been increasing its revenues. ");


            }
            else if (revenue[0] == "unclear")
            {
                aux.Add("Revenues do not have a clear trend. ");
            }
            else
            {
                aux.Add("Revenues have been decreasing. It is important to understand why and whether or not this trend will continue. ");
            }

            if (revenue[1] == "high volatility")
            {
                aux.Add("Revenues are showing a high volatility in the past. What is the reason behind this? ");
            }
            else
            {

            }

            var stringcomment = string.Join("", aux);

            return stringcomment;
        }

        public string commentsNOPLAT(List<string> revenue)
        {
            List<string> aux = new List<string>(); ;

            if (revenue[0] == "increase")
            {
                aux.Add("The company has been increasing its NOPLAT. ");


            }
            else if (revenue[0] == "unclear")
            {
                aux.Add("Revenues do not have a clear trend. ");
            }
            else
            {
                aux.Add("NOPLAT have been decreasing. It is important to understand why and whether or not this trend will continue. ");
            }

            if (revenue[1] == "high volatility")
            {
                aux.Add("NOPLAT is showing a high volatility in the past. What is the reason behind this? ");
            }
            else
            {

            }

            var stringcomment = string.Join("", aux);

            return stringcomment;
        }
        public string commentsIC(List<string> revenue)
        {
            List<string> aux = new List<string>(); ;

            if (revenue[0] == "increase")
            {
                aux.Add("The company has been increasing its Invested Capital. ");


            }
            else if (revenue[0] == "unclear")
            {
                aux.Add("Invested Capital does not have a clear trend. ");
            }
            else
            {
                aux.Add("Invested Capital have been decreasing. It is important to understand why and whether or not this trend will continue. ");
            }


            var stringcomment = string.Join("", aux);

            return stringcomment;
        }
    }

    
}
