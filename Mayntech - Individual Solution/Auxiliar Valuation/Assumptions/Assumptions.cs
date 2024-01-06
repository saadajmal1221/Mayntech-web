using Mayntech___Individual_Solution.Auxiliar.Auxiliar.FMP;
using Mayntech___Individual_Solution.Auxiliar.Auxiliar.Other;
using Mayntech___Individual_Solution.Auxiliar_Analysis.Auxiliar.Other;
using Mayntech___Individual_Solution.Auxiliar_Valuation.Financials_projections;
using Mayntech___Individual_Solution.Pages.Solutions;
using OfficeOpenXml;
using OfficeOpenXml.Style;
using System.Collections.Generic;
using System.Data.Common;
using System.Drawing;
using System.Globalization;


namespace Mayntech___Individual_Solution.Auxiliar_Valuation.Assumptions
{
    public class Assumptions
    {
        public async Task assumptionsBuilderAsync(ExcelPackage package, List<FinancialStatements> incomeStatement, int numberOfYears, 
            CompanyProfile companyProfile, List<MarketRiskPremium> marketRiskPremia, List<CompanyNotes> companyNotes, Taxes tax, 
            WaccDamodaran WaccInputs, int quarters)

        {
            ExcelNextCol columnName = new ExcelNextCol();
            var worksheet = package.Workbook.Worksheets.Add("Assumptions");
            worksheet.View.ShowGridLines = false;
            worksheet.View.ZoomScale = 80;
            int row = 2;
            int col = 2;

            //Result : [Beta], [MarketRiskPremium], [riskFree], [costOfDebt], [costOfequity], [countryRiskPremium]
            List<double> WaccCalc = await costOfCapitalCalcAsync(companyProfile, marketRiskPremia, companyNotes);

            for (int i = 0; i < 4; i++)
            {
                if (i==0)
                {
                    worksheet.Cells[row, col].Value = "Assumptions Valuation";
                    worksheet.Cells[row, col].Style.Font.Bold = true;
                    worksheet.Cells[row, col].Style.Font.Color.SetColor(Color.White);

                    worksheet.Cells[row + 3, col+i].Value = "Revenue";
                    worksheet.Cells[row + 4, col + i].Value = "Cost of revenues";
                    worksheet.Cells[row + 5, col + i].Value = "Operating costs";
                    worksheet.Cells[row + 6, col + i].Value = "Change in PP&E";
                    worksheet.Cells[row + 7, col + i].Value = "Change in Intangibles";
                    worksheet.Cells[row + 8, col + i].Value = "Change in NWC";





                    worksheet.Cells[row + 1, col + i].Value = "Description";
                    worksheet.Cells[row + 1, col + i].Style.Font.Color.SetColor(Cores.CorTexto);
                    worksheet.Cells[row + 1, col + i].Style.Border.Bottom.Style = ExcelBorderStyle.Thin;
                    worksheet.Cells[row + 1, col + i].Style.Font.Bold = true;

                    worksheet.Columns[col + i].Width = 20;

                }

                worksheet.Cells[row, col + i].Style.Fill.PatternType = ExcelFillStyle.Solid;
                worksheet.Cells[row, col + i].Style.Fill.BackgroundColor.SetColor(Cores.CorPrincipal);

                int LastYear = incomeStatement[incomeStatement.Count()-1].Date.Year;
                int NextYear = LastYear + 1;

                int AfterFiveYears = NextYear + 5;
                int LastProjectedYear = AfterFiveYears + 3;

                if (i==1)
                {
                    for (int a = 1; a < 6; a++)
                    {
                        worksheet.Cells[row + 3 + a, col + i].Style.Fill.PatternType = ExcelFillStyle.Solid;
                        worksheet.Cells[row + 3 + a, col + i].Style.Fill.BackgroundColor.SetColor(Cores.CorInput);
                    }

                    worksheet.Cells[row + 3, col + i].Style.Fill.PatternType = ExcelFillStyle.Solid;
                    worksheet.Cells[row + 3, col + i].Style.Fill.BackgroundColor.SetColor(Cores.CorCinzentoBackground);

                    worksheet.Cells[row + 1, col + i].Value = LastYear + " - " + NextYear + " growth";
                    worksheet.Cells[row + 1, col + i].Style.HorizontalAlignment = ExcelHorizontalAlignment.Right;
                    worksheet.Cells[row + 1, col + i].Style.Font.Color.SetColor(Cores.CorTexto);
                    worksheet.Cells[row + 1, col + i].Style.Border.Bottom.Style = ExcelBorderStyle.Thin;
                    worksheet.Cells[row + 1, col + i].Style.Font.Bold = true;

                    string columnQuarter = columnName.GetExcelColumnName(4 + numberOfYears);
                    string columnLastQuarter = columnName.GetExcelColumnName(3 + numberOfYears + quarters);
                    string columnLastYear = columnName.GetExcelColumnName(2 + numberOfYears);
                    string columnPenultimoYear = columnName.GetExcelColumnName(1 + numberOfYears);
                    string columnantePenultiYear = columnName.GetExcelColumnName(numberOfYears);
                    string columnFirstYear = columnName.GetExcelColumnName(3);


                    //Revenue
                    if (quarters > 0)
                    {

                        if (quarters ==1)
                        {
                            string quartersAssumption = "(SUM('P&L'!" + columnQuarter + "5:'P&L'!" + columnLastQuarter + "5)*4)/('P&L'!" + columnLastYear + "5)-1";
                            WriteFormula(worksheet, row, col + i, columnPenultimoYear, columnLastYear, quartersAssumption, 6, quarters);
                        }
                        else if (quarters == 2)
                        {
                            string quartersAssumption = "(SUM('P&L'!" + columnQuarter + "5:'P&L'!" + columnLastQuarter + "5)*2)/('P&L'!" + columnLastYear + "5)-1";
                            WriteFormula(worksheet, row, col + i, columnPenultimoYear, columnLastYear, quartersAssumption, 6, quarters);
                        }
                        else if (quarters == 3)
                        {
                            string quartersAssumption = "(SUM('P&L'!" + columnQuarter + "5:'P&L'!" + columnLastQuarter + "5)*(4 / 3))/('P&L'!" + columnLastYear + "5)-1";
                            WriteFormula(worksheet, row, col + i, columnPenultimoYear, columnLastYear, quartersAssumption, 6, quarters);
                        }
                        
                    }
                    else
                    {
                        worksheet.Cells[row + 3, col + i].Value = 0;
                    }

                    
                    worksheet.Cells[row + 3, col + i].Style.Numberformat.Format = "0.00%";

                    worksheet.Cells[row + 4, col + i].Value = 0;
                    worksheet.Cells[row + 4, col + i].Style.Numberformat.Format = "0.00%";

                    worksheet.Cells[row + 5, col + i].Value = 0;
                    worksheet.Cells[row + 5, col + i].Style.Numberformat.Format = "0.00%";


                    worksheet.Cells[row + 6, col + i].Value = 0;
                    worksheet.Cells[row + 6, col + i].Style.Numberformat.Format = "0.00%";

                    worksheet.Cells[row + 7, col + i].Value = 0;
                    worksheet.Cells[row + 7, col + i].Style.Numberformat.Format = "0.00%";

                    worksheet.Cells[row + 8, col + i].Value = 0;
                    worksheet.Cells[row + 8, col + i].Style.Numberformat.Format = "0.00%";

                    worksheet.Columns[col + i].Width = 20;
                }
                if (i==2)
                {
                    for (int a = 0; a < 6; a++)
                    {
                        worksheet.Cells[row + 3 + a, col + i].Style.Fill.PatternType = ExcelFillStyle.Solid;
                        worksheet.Cells[row + 3 + a, col + i].Style.Fill.BackgroundColor.SetColor(Cores.CorInput);
                    }
                    worksheet.Cells[row + 1, col + i].Value = NextYear + " - " + AfterFiveYears + " (CAGR)";
                    worksheet.Cells[row + 1, col + i].Style.HorizontalAlignment = ExcelHorizontalAlignment.Right;
                    worksheet.Cells[row + 1, col + i].Style.Font.Color.SetColor(Cores.CorTexto);
                    worksheet.Cells[row + 1, col + i].Style.Border.Bottom.Style = ExcelBorderStyle.Thin;
                    worksheet.Cells[row + 1, col + i].Style.Font.Bold = true;


                    worksheet.Cells[row + 3, col + i].Value = 0;
                    worksheet.Cells[row + 3, col + i].Style.Numberformat.Format = "0.00%";

                    worksheet.Cells[row + 4, col + i].Value = 0;
                    worksheet.Cells[row + 4, col + i].Style.Numberformat.Format = "0.00%";

                    worksheet.Cells[row + 5, col + i].Value = 0;
                    worksheet.Cells[row + 5, col + i].Style.Numberformat.Format = "0.00%";

                    worksheet.Cells[row + 6, col + i].Value = 0;
                    worksheet.Cells[row + 6, col + i].Style.Numberformat.Format = "0.00%";

                    worksheet.Cells[row + 7, col + i].Value = 0;
                    worksheet.Cells[row + 7, col + i].Style.Numberformat.Format = "0.00%";

                    worksheet.Cells[row + 8, col + i].Value = 0;
                    worksheet.Cells[row + 8, col + i].Style.Numberformat.Format = "0.00%";


                    worksheet.Columns[col + i].Width = 20;
                }
                if (i == 3)
                {
                    for (int a = 0; a < 6; a++)
                    {
                        worksheet.Cells[row+3+a, col +i].Style.Fill.PatternType = ExcelFillStyle.Solid;
                        worksheet.Cells[row + 3 + a, col +i].Style.Fill.BackgroundColor.SetColor(Cores.CorInput);
                    }

                    worksheet.Cells[row + 1, col + i].Value = AfterFiveYears + " - " + LastProjectedYear + " (CAGR)";
                    worksheet.Cells[row + 1, col + i].Style.HorizontalAlignment = ExcelHorizontalAlignment.Right;
                    worksheet.Cells[row + 1, col + i].Style.Font.Color.SetColor(Cores.CorTexto);
                    worksheet.Cells[row + 1, col + i].Style.Border.Bottom.Style = ExcelBorderStyle.Thin;
                    worksheet.Cells[row + 1, col + i].Style.Font.Bold = true;


                    worksheet.Cells[row + 3, col + i].Value = 0;
                    worksheet.Cells[row + 3, col + i].Style.Numberformat.Format = "0.00%";

                    worksheet.Cells[row + 4, col + i].Value = 0;
                    worksheet.Cells[row + 4, col + i].Style.Numberformat.Format = "0.00%";

                    worksheet.Cells[row + 5, col + i].Value = 0;
                    worksheet.Cells[row + 5, col + i].Style.Numberformat.Format = "0.00%";

                    worksheet.Cells[row + 6, col + i].Value = 0;
                    worksheet.Cells[row + 6, col + i].Style.Numberformat.Format = "0.00%";

                    worksheet.Cells[row + 7, col + i].Value = 0;
                    worksheet.Cells[row + 7, col + i].Style.Numberformat.Format = "0.00%";

                    worksheet.Cells[row + 8, col + i].Value = 0;
                    worksheet.Cells[row + 8, col + i].Style.Numberformat.Format = "0.00%";


                    worksheet.Columns[col + i].Width = 20;
                }
            }

            for (int i = 5; i < 7; i++)
            {
                if (i==5)
                {
                    worksheet.Cells[row, col+i].Value = "Other Assumptions";
                    worksheet.Cells[row, col+i].Style.Font.Bold = true;
                    worksheet.Cells[row, col + i].Style.Font.Color.SetColor(Color.White);

                    worksheet.Cells[row + 1, col + i].Value = "Description";
                    worksheet.Cells[row + 1, col + i].Style.Font.Color.SetColor(Cores.CorTexto);
                    worksheet.Cells[row + 1, col + i].Style.Border.Bottom.Style = ExcelBorderStyle.Thin;
                    worksheet.Cells[row + 1, col + i].Style.Font.Bold = true;


                    worksheet.Cells[row + 3, col + i].Value = "WACC";
                    worksheet.Cells[row + 4, col + i].Value = "Tax Rate";
                    worksheet.Cells[row + 5, col + i].Value = "Operating cash";
                    worksheet.Cells[row + 6, col + i].Value = "Depreciation & amortization rate";
                    worksheet.Cells[row + 7, col + i].Value = "Market Risk Premium";
                    worksheet.Cells[row + 8, col + i].Value = "Country Risk Premium";
                    worksheet.Cells[row + 9, col + i].Value = "Company Beta";
                    worksheet.Cells[row + 10, col + i].Value = "Risk Free";
                    worksheet.Cells[row + 11, col + i].Value = "Cost of Debt";
                    worksheet.Cells[row + 12, col + i].Value = "Cost of Equity";
                    worksheet.Cells[row + 13, col + i].Value = "Debt to Capital";

                    if (WaccInputs !=null)
                    {
                        worksheet.Cells[row + 15, col + i].Value = "Industry Debt to Capital";
                    }
                    

                    worksheet.Columns[col + i].Width = 30;
                }


                if (i==6)
                {
                    string column = columnName.GetExcelColumnName(2 + numberOfYears);
                    string columnGrowthRate = columnName.GetExcelColumnName(2 + numberOfYears + 12);

                    worksheet.Cells[row + 1, col + i].Value = "Value";
                    worksheet.Cells[row + 1, col + i].Style.HorizontalAlignment = ExcelHorizontalAlignment.Right;
                    worksheet.Cells[row + 1, col + i].Style.Font.Color.SetColor(Cores.CorTexto);
                    worksheet.Cells[row + 1, col + i].Style.Border.Bottom.Style = ExcelBorderStyle.Thin;
                    worksheet.Cells[row + 1, col + i].Style.Font.Bold = true;

                    worksheet.Rows[row + 2].Height = 3;

                    worksheet.Cells[row + 3, col + i].Formula = "=(1-H6)*H13*H15 + H14*(1-H15)";
                    worksheet.Cells[row + 3, col + i].Style.Numberformat.Format = "0.00%";

                    //Tax
                    NumberFormatInfo provider = new NumberFormatInfo();
                    provider.NumberDecimalSeparator = ".";

                    double beta = new double();
                    double costOfDebt = new double();
                    double debtToCapital = new double();

                    if (WaccInputs !=null)
                    {
                        beta = double.Parse(WaccInputs.Beta, provider);
                        costOfDebt = double.Parse(WaccInputs.CostofDebt, provider);
                        debtToCapital = double.Parse(WaccInputs.debtToCapital, provider);
                    }
                    double output = double.Parse(tax.year2022, provider);


                    worksheet.Cells[row + 4, col + i].Value = output/100;
                    worksheet.Cells[row + 4, col + i].Style.Numberformat.Format = "0.00%";

                    worksheet.Cells[row + 5, col + i].Formula = "=IF(AVERAGE(Auxiliar!C8:" + column + "8)<0,0.02,0.05)";
                    worksheet.Cells[row + 5, col + i].Style.Numberformat.Format = "0.00%";

                    worksheet.Cells[row + 6, col + i].Formula = "=IF(AVERAGE(Auxiliar!C14:" + column + "14)<0.5, AVERAGE(Auxiliar!C14:" + column + "14), 0.5)";
                    worksheet.Cells[row + 6, col + i].Style.Numberformat.Format = "0.00%";

                    //Market Risk Premium
                    worksheet.Cells[row + 7, col + i].Value = ((WaccCalc[1]) / 100);
                    worksheet.Cells[row + 7, col + i].Style.Numberformat.Format = "0.00%";

                    //Country Risk Premium
                    worksheet.Cells[row + 8, col + i].Value = WaccCalc[5] / 100;
                    worksheet.Cells[row + 8, col + i].Style.Numberformat.Format = "0.00%";

                    //Beta
                    worksheet.Cells[row + 9, col + i].Value = WaccCalc[0];
                    worksheet.Cells[row + 9, col + i].Style.Numberformat.Format = "#,##0.00;(#,##0.00);-";

                    //Risk Free
                    worksheet.Cells[row + 10, col + i].Value = WaccCalc[2]/100;
                    worksheet.Cells[row + 10, col + i].Style.Numberformat.Format = "0.00%";

                    //Cost Of Debt
                    if (WaccInputs != null)
                    {
                        worksheet.Cells[row + 11, col + i].Value = ((WaccCalc[3] / 100) + (costOfDebt / 100)) / 2;
                        worksheet.Cells[row + 11, col + i].Style.Numberformat.Format = "0.00%";
                    }
                    else
                    {
                        worksheet.Cells[row + 11, col + i].Value = (WaccCalc[3] / 100) ;
                        worksheet.Cells[row + 11, col + i].Style.Numberformat.Format = "0.00%";
                    }

                    //Cost Of equity
                    worksheet.Cells[row + 12, col + i].Formula = "=H12+H11*H9+H10";
                    worksheet.Cells[row + 12, col + i].Style.Numberformat.Format = "0.00%";

                    if (WaccInputs != null)
                    {
                        //capital Structure
                        worksheet.Cells[row + 13, col + i].Formula = "=(Auxiliar!C39 + H17)/2";
                        worksheet.Cells[row + 13, col + i].Style.Numberformat.Format = "0.00%";
                    }
                    else
                    {
                        //capital Structure
                        worksheet.Cells[row + 13, col + i].Formula = "=Auxiliar!C39";
                        worksheet.Cells[row + 13, col + i].Style.Numberformat.Format = "0.00%";
                    }

                    if (WaccInputs != null)
                    {
                        //Industry debt to capital rate
                        worksheet.Cells[row + 15, col + i].Value = debtToCapital / 100;
                        worksheet.Cells[row + 15, col + i].Style.Numberformat.Format = "0.00%";
                    }
                    else
                    {

                    }


                }

                worksheet.Cells[row, col + i].Style.Fill.PatternType = ExcelFillStyle.Solid;
                worksheet.Cells[row, col + i].Style.Fill.BackgroundColor.SetColor(Cores.CorPrincipal);
            }

            for (int i = 10; i < 12; i++)
            {
                if (i == 10)
                {
                    worksheet.Cells[row, col + i].Value = "Terminal Value Assumptions";
                    worksheet.Cells[row, col + i].Style.Font.Bold = true;
                    worksheet.Cells[row, col + i].Style.Font.Color.SetColor(Color.White);

                    worksheet.Cells[row + 1, col + i].Value = "Description";
                    worksheet.Cells[row + 1, col + i].Style.Font.Color.SetColor(Cores.CorTexto);
                    worksheet.Cells[row + 1, col + i].Style.Border.Bottom.Style = ExcelBorderStyle.Thin;
                    worksheet.Cells[row + 1, col + i].Style.Font.Bold = true;


                    worksheet.Cells[row + 3, col + i].Value = "Inflation Rate";
                    worksheet.Cells[row + 4, col + i].Value = "WACC";
                    worksheet.Cells[row + 5, col + i].Value = "Growth Rate";


                    worksheet.Columns[col + i].Width = 30;
                }


                if (i == 11)
                {
                    string column = columnName.GetExcelColumnName(2 + numberOfYears);
                    string columnGrowthRate = columnName.GetExcelColumnName(2 + numberOfYears + 12);
                    string LastColumn = columnName.GetExcelColumnName(2 + numberOfYears + 9);
                    string FirstProjection = columnName.GetExcelColumnName(2 + numberOfYears + 9);

                    worksheet.Cells[row + 1, col + i].Value = "Value";
                    worksheet.Cells[row + 1, col + i].Style.Font.Color.SetColor(Cores.CorTexto);
                    worksheet.Cells[row + 1, col + i].Style.HorizontalAlignment = ExcelHorizontalAlignment.Right;
                    worksheet.Cells[row + 1, col + i].Style.Border.Bottom.Style = ExcelBorderStyle.Thin;
                    worksheet.Cells[row + 1, col + i].Style.Font.Bold = true;

                    worksheet.Rows[row + 2].Height = 3;

                    worksheet.Cells[row + 3, col + i].Formula = "=0.02";
                    worksheet.Cells[row + 3, col + i].Style.Numberformat.Format = "0.00%";


                    worksheet.Cells[row + 4, col + i].Formula = "=H5";
                    worksheet.Cells[row + 4, col + i].Style.Numberformat.Format = "0.00%";

                    worksheet.Cells[row + 5, col + i].Formula = "=IF('Growth Analysis'!" + columnGrowthRate + "12>0.035, 0.035, IF('Growth Analysis'!" + columnGrowthRate+ "12<0,0,'Growth Analysis'!" + columnGrowthRate+ "12)) ";
                    worksheet.Cells[row + 5, col + i].Style.Numberformat.Format = "0.00%";



                }

                worksheet.Cells[row, col + i].Style.Fill.PatternType = ExcelFillStyle.Solid;
                worksheet.Cells[row, col + i].Style.Fill.BackgroundColor.SetColor(Cores.CorPrincipal);
            }

        }

        public void WriteFormula(ExcelWorksheet worksheet, int row, int col, string columnPenultimoYear, string columnLastYear, string quartersAssumption, int rowCaption, int quarters)
        {
            if (quarters>0)
            {
                string formula = "=" + quartersAssumption;
                worksheet.Cells[row + 3, col].Formula = formula;
            }
            
        }

        public async Task<List<double>> costOfCapitalCalcAsync(CompanyProfile companyProfile, List<MarketRiskPremium> marketRiskPremia,
            List<CompanyNotes> companyNotes)
        {
            //Result : [Beta], [MarketRiskPremium], [riskFree], [costOfDebt], [costOfequity], [countryRiskPremium]
            List<double> result = new List<double>();
            List<TreasuryRates> treasuryRates = new List<TreasuryRates>();

            //double costOfDebt = CostOfDebt(companyNotes);

            double beta = (double)companyProfile.profile.beta;
            result.Add(beta);
            string countryaux = companyProfile.profile.country;

            RegionInfo countriesInfo = new RegionInfo(countryaux);
            string country = countriesInfo.EnglishName;

            double costOfDebt = 0;
            double marketriskpremium = 0;
            double countryRiskPremium = 0;

            MarketRiskPremium? marketRisk = marketRiskPremia.Where(i => i.country == country).FirstOrDefault();
            if (marketRisk != null)
            {
                marketriskpremium = marketRisk.totalEquityRiskPremium;
                countryRiskPremium = marketRisk.countryRiskPremium;
                result.Add(marketriskpremium);
            }
            else
            {
                result.Add(10);
            }

            FinancialMP fmp = new FinancialMP();

            int year = DateTime.Now.Year;
            int month = DateTime.Now.Month;
            int day = DateTime.Now.Day;
            int dayMinusTen = DateTime.Now.Day;
            int yearMinusOne = DateTime.Now.Year;
            int monthMinusOne = DateTime.Now.Month;

            if (day -10 <0)
            {
                dayMinusTen = 25;
                
                if (month-1<=0)
                {
                    monthMinusOne = 12;
                    yearMinusOne -= 1;
                }
                else
                {
                    monthMinusOne -= 1;
                }
            }
            else
            {
                dayMinusTen = day -10;
            }

            

            string APITreasuryrates = "https://financialmodelingprep.com/api/v4/treasury?from="+yearMinusOne +"-" + monthMinusOne+ "-"+ dayMinusTen+ "&to=" + year +"-" + month+"-"+ day+ "&apikey=d7ba591715e23c3d2ac3b0aec6dea138";
            treasuryRates = await fmp.GetTreasuryRates(APITreasuryrates);

            double riskFree = (double)treasuryRates[0].year10;
            string date = treasuryRates[0].date;
            result.Add(riskFree);

            if (beta >1.7)
            {
                costOfDebt = 6 + riskFree;
            }
            else if (beta > 1.4)
            {
                costOfDebt = 4 + riskFree;
            }
            else if (beta > 1)
            {
                costOfDebt = 3 + riskFree;
            }
            else
            {
                costOfDebt = 2 + riskFree;
            }
            result.Add(costOfDebt);

            double CostOfEquity = riskFree + beta* marketriskpremium + countryRiskPremium;

            result.Add(CostOfEquity);
            result.Add(countryRiskPremium);


            return result;
        }

        public double CostOfDebt(List<CompanyNotes> companyNotes)
        {
            foreach (var CompanyNote in companyNotes)
            {
                char[] notesDue = CompanyNote.title.ToCharArray();

                List<char> cost = new List<char>();
                List<char> year = new List<char>();

                string costAux = null;
                string yearAux = null;

                foreach (var letter in notesDue)
                {


                    if (letter<5)
                    {
                        cost.Add(notesDue[letter]);
                        costAux.Append(notesDue[letter]);
                    }
                    if (letter>notesDue.Count()-5)
                    {
                        year.Add(notesDue[letter]);
                        yearAux.Append(notesDue[letter]);
                    }

                }

                
                
            }
            return 0;
        }
    }
}
