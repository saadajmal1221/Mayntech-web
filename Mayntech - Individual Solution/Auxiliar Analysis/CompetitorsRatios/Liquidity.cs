using Mayntech___Individual_Solution.Auxiliar.Auxiliar.FMP;
using Mayntech___Individual_Solution.Pages;
using OfficeOpenXml;
using OfficeOpenXml.Style;
using System.Drawing;
using System.Reflection;

namespace Mayntech___Individual_Solution.Auxiliar_Analysis.CompetitorsRatios
{
    public class Liquidity : Ratios
    {
        public void LiquidityConstruction(ExcelPackage package, int numberOfYears, List<string> comments,
            string companyTick, int numberOfYearsIncomeStatement)
        {


            List<double> points = new List<double>();

            var worksheet = package.Workbook.Worksheets.Add("Liquidity");
            worksheet.View.ShowGridLines = false;
            worksheet.View.ZoomScale = 80;
            int row = 9;
            int col = 2;

            for (int i = 0; i < 5; i++)
            {
                worksheet.Rows[2 + i].OutlineLevel = 1;
                worksheet.Rows[2 + i].Collapsed = true;
            }

            worksheet.Row(7).Style.Fill.PatternType = ExcelFillStyle.Solid;
            worksheet.Row(7).Style.Fill.BackgroundColor.SetColor(Color.Gray);
            worksheet.Cells[7, 2].Value = "Ratios";
            worksheet.Cells[7, 2].Style.Font.Bold = true;
            worksheet.Cells[7, 2].Style.Font.Color.SetColor(Color.White);

            int numberOfcolumns = Math.Min(5, numberOfYears);


            //current Ratio
            try
            {
                List<double> ReferenceCompany = new List<double>();
                IDictionary<string, List<double>> AllCompanyRatios = new Dictionary<string, List<double>>();
                IDictionary<string, double> AllCompaniesAverage = new Dictionary<string, double>();
                IDictionary<string, double> AllCompaniesSd = new Dictionary<string, double>();
                //Adiciona os dados da empresa referencia
                for (int i = 0; i < numberOfcolumns; i++)
                {
                    try
                    {
                        double CompanyAnalysedRatio = SolutionModel.balances[(numberOfcolumns-1) - i].TotalCurrentAssets / SolutionModel.balances[(numberOfcolumns - 1) - i].totalCurrentLiabilities;
                        ReferenceCompany.Add(CompanyAnalysedRatio);
                    }
                    catch (Exception)
                    {

                    }

                }

                AllCompanyRatios.Add(companyTick, ReferenceCompany);
                AllCompaniesAverage.Add(companyTick, ReferenceCompany.Average());
                AllCompaniesSd.Add(companyTick, StandardDeviation(ReferenceCompany));




                //Adiciona os dados de todas as empresas
                foreach (KeyValuePair<string, List<FinancialStatements>> item in SolutionModel.BalanceSheetDict)
                {
                    List<double> individualCompanyRatios = new List<double>();

                    try
                    {
                        if (item.Value.Count() >= Math.Min(5, numberOfYears) && item.Value[0].totalCurrentLiabilities > 0)
                        {

                            try
                            {
                                for (int i = 0; i < numberOfcolumns; i++)
                                {
                                    double ratio = item.Value[i].TotalCurrentAssets / item.Value[i].totalCurrentLiabilities;
                                    individualCompanyRatios.Add(ratio);
                                }
                            }
                            catch (Exception)
                            {

                                continue;
                            }

                            //Adiciona os ratios, média e Sd da empresa à lista
                            AllCompanyRatios.Add(item.Key, individualCompanyRatios);
                            AllCompaniesAverage.Add(item.Key, individualCompanyRatios.Average());
                            AllCompaniesSd.Add(item.Key, StandardDeviation(individualCompanyRatios));
                        }
                    }
                    catch (Exception)
                    {

                        continue;
                    }

                    
                }

                //Primeiro comentário é para os casos opostos do suposto, o segundo é para outlier negativo e o terceiro
                //(cont) é para outlier positivo
                List<string> commentsRatio = new List<string> {"Current Ratio is below the average of its competitors", 
                    "Current ratio is extremely low compared to its peers", 
                "Current Ratio is extremely high compared to its peers"};

                points.Add(PointCalculator(AllCompaniesAverage, AllCompaniesSd, companyTick, 1, commentsRatio , "Current Ratio"));

                RatioConst(worksheet, "Current Ratio", "='BS'!", "12 / 'BS'!", "32", null, null, col, row, numberOfYears, AllCompanyRatios, AllCompaniesAverage, companyTick, numberOfYearsIncomeStatement,0);
                row += AllCompanyRatios.Count() + 4;
            }
            catch (Exception)
            {

                throw;
            }







            //Quick Ratio
            try
            {
                List<double> ReferenceCompany = new List<double>();
                IDictionary<string, List<double>> AllCompanyRatios = new Dictionary<string, List<double>>();
                IDictionary<string, double> AllCompaniesAverage = new Dictionary<string, double>();
                IDictionary<string, double> AllCompaniesSd = new Dictionary<string, double>();
                //Adiciona os dados da empresa referencia
                for (int i = 0; i < numberOfcolumns; i++)
                {
                    try
                    {
                        double CompanyAnalysedRatio = (SolutionModel.balances[(numberOfcolumns - 1) - i].CashAndCashEquivalents + SolutionModel.balances[(numberOfcolumns - 1) - i].ShortTermInvestments + SolutionModel.balances[(numberOfcolumns - 1) - i].NetReceivables) / SolutionModel.balances[(numberOfcolumns - 1) - i].totalCurrentLiabilities;
                        ReferenceCompany.Add(CompanyAnalysedRatio);
                    }
                    catch (Exception)
                    {

                        
                    }

                }
                AllCompanyRatios.Add(companyTick, ReferenceCompany);
                AllCompaniesAverage.Add(companyTick, ReferenceCompany.Average());
                AllCompaniesSd.Add(companyTick, StandardDeviation(ReferenceCompany));


                //Adiciona os dados de todas as empresas
                foreach (KeyValuePair<string, List<FinancialStatements>> item in SolutionModel.BalanceSheetDict)
                {
                    List<double> individualCompanyRatios = new List<double>();

                    try
                    {
                        if (item.Value.Count() >= Math.Min(5, numberOfYears) && item.Value[0].totalCurrentLiabilities > 0)
                        {

                            try
                            {
                                for (int i = 0; i < numberOfcolumns; i++)
                                {
                                    double ratio = (item.Value[i].CashAndCashEquivalents + item.Value[i].ShortTermInvestments + item.Value[i].NetReceivables) / item.Value[i].totalCurrentLiabilities;
                                    individualCompanyRatios.Add(ratio);
                                }
                            }
                            catch (Exception)
                            {

                                continue;
                            }

                            //Adiciona os ratios, média e Sd da empresa à lista
                            AllCompanyRatios.Add(item.Key, individualCompanyRatios);
                            AllCompaniesAverage.Add(item.Key, individualCompanyRatios.Average());
                            AllCompaniesSd.Add(item.Key, StandardDeviation(individualCompanyRatios));
                        }
                    }
                    catch (Exception)
                    {

                        continue;
                    }
                    
                }


                //Primeiro comentário é para os casos opostos do suposto, o segundo é para outlier negativo e o terceiro
                //(cont) é para outlier positivo
                List<string> commentsRatio = new List<string> {"Current Ratio is below the average of its competitors",
                    "Current ratio is extremely low compared to its peers",
                "Current Ratio is extremely high compared to its peers"};

                points.Add(PointCalculator(AllCompaniesAverage, AllCompaniesSd, companyTick, 1, commentsRatio, "Quick ratio"));

                RatioConst(worksheet, "Quick Ratio", "=('BS'!", "7 + 'BS'!", "8+'BS'!", "9)/'BS'!", "32", col, row, numberOfYears, AllCompanyRatios, AllCompaniesAverage,companyTick, numberOfYearsIncomeStatement,0);
                row += AllCompanyRatios.Count() + 4;
            }
            catch (Exception)
            {

                throw;
            }



            //Cash Ratio
            try
            {
                List<double> ReferenceCompany = new List<double>();
                IDictionary<string, List<double>> AllCompanyRatios = new Dictionary<string, List<double>>();
                IDictionary<string, double> AllCompaniesAverage = new Dictionary<string, double>();
                IDictionary<string, double> AllCompaniesSd = new Dictionary<string, double>();
                //Adiciona os dados da empresa referencia
                for (int i = 0; i < numberOfcolumns; i++)
                {
                    try
                    {
                        double CompanyAnalysedRatio = (SolutionModel.balances[(numberOfcolumns-1) - i].CashAndCashEquivalents + SolutionModel.balances[(numberOfcolumns - 1) - i].ShortTermInvestments) / SolutionModel.balances[(numberOfcolumns - 1) - i].totalCurrentLiabilities;
                        ReferenceCompany.Add(CompanyAnalysedRatio);
                    }
                    catch (Exception)
                    {


                    }

                }
                AllCompanyRatios.Add(companyTick, ReferenceCompany);
                AllCompaniesAverage.Add(companyTick, ReferenceCompany.Average());
                AllCompaniesSd.Add(companyTick, StandardDeviation(ReferenceCompany));


                //Adiciona os dados de todas as empresas
                foreach (KeyValuePair<string, List<FinancialStatements>> item in SolutionModel.BalanceSheetDict)
                {
                    List<double> individualCompanyRatios = new List<double>();

                    try
                    {
                        if (item.Value.Count() >= Math.Min(5, numberOfYears) && item.Value[0].totalCurrentLiabilities > 0)
                        {

                            try
                            {
                                for (int i = 0; i < numberOfcolumns; i++)
                                {
                                    double ratio = (item.Value[i].CashAndCashEquivalents + item.Value[i].ShortTermInvestments) / item.Value[i].totalCurrentLiabilities;
                                    individualCompanyRatios.Add(ratio);
                                }
                            }
                            catch (Exception)
                            {

                                continue;
                            }

                            //Adiciona os ratios, média e Sd da empresa à lista
                            AllCompanyRatios.Add(item.Key, individualCompanyRatios);
                            AllCompaniesAverage.Add(item.Key, individualCompanyRatios.Average());
                            AllCompaniesSd.Add(item.Key, StandardDeviation(individualCompanyRatios));
                        }
                    }
                    catch (Exception)
                    {

                        continue;
                    }

                }


                //Primeiro comentário é para os casos opostos do suposto, o segundo é para outlier negativo e o terceiro
                //(cont) é para outlier positivo
                List<string> commentsRatio = new List<string> {"Cash Ratio is below the average of its competitors",
                    "Cash ratio is extremely low compared to its peers",
                "Cash Ratio is extremely high compared to its peers"};

                points.Add(PointCalculator(AllCompaniesAverage, AllCompaniesSd, companyTick, 1, commentsRatio, "Cash ratio"));

                RatioConst(worksheet, "Cash ratio", "=('BS'!", "7 + 'BS'!", "8)/'BS'!", "32", null, col, row, numberOfYears, AllCompanyRatios, AllCompaniesAverage, companyTick, numberOfYearsIncomeStatement,0);
                row += AllCompanyRatios.Count() + 4;
            }
            catch (Exception)
            {

                throw;
            }





            //Operating cash flow ratio
            try
            {
                List<double> ReferenceCompany = new List<double>();
                IDictionary<string, List<double>> AllCompanyRatios = new Dictionary<string, List<double>>();
                IDictionary<string, double> AllCompaniesAverage = new Dictionary<string, double>();
                IDictionary<string, double> AllCompaniesSd = new Dictionary<string, double>();
                //Adiciona os dados da empresa referencia
                for (int i = 0; i < numberOfcolumns; i++)
                {
                    try
                    {
                        double CompanyAnalysedRatio = (SolutionModel.cashFlow[(numberOfcolumns-1) - i].NetCashProvidedByOperatingActivities) / SolutionModel.balances[(numberOfcolumns - 1) - i].totalCurrentLiabilities;
                        ReferenceCompany.Add(CompanyAnalysedRatio);
                    }
                    catch (Exception)
                    {


                    }

                }
                AllCompanyRatios.Add(companyTick, ReferenceCompany);
                AllCompaniesAverage.Add(companyTick, ReferenceCompany.Average());
                AllCompaniesSd.Add(companyTick, StandardDeviation(ReferenceCompany));


                //Adiciona os dados de todas as empresas
                foreach (KeyValuePair<string, List<FinancialStatements>> item in SolutionModel.BalanceSheetDict)
                {
                    List<double> individualCompanyRatios = new List<double>();

                    try
                    {
                        if (item.Value.Count() >= Math.Min(5, numberOfYears) && item.Value[0].totalCurrentLiabilities > 0)
                        {

                            try
                            {
                                for (int i = 0; i < numberOfcolumns; i++)
                                {
                                    double ratio = (item.Value[i].NetCashProvidedByOperatingActivities) / item.Value[i].totalCurrentLiabilities;
                                    individualCompanyRatios.Add(ratio);
                                }
                            }
                            catch (Exception)
                            {

                                continue;
                            }

                            //Adiciona os ratios, média e Sd da empresa à lista
                            AllCompanyRatios.Add(item.Key, individualCompanyRatios);
                            AllCompaniesAverage.Add(item.Key, individualCompanyRatios.Average());
                            AllCompaniesSd.Add(item.Key, StandardDeviation(individualCompanyRatios));
                        }
                    }
                    catch (Exception)
                    {

                        continue;
                    }

                }


                //Primeiro comentário é para os casos opostos do suposto, o segundo é para outlier negativo e o terceiro
                //(cont) é para outlier positivo
                List<string> commentsRatio = new List<string> {"Operating Cash flow ratio is below the average of its competitors",
                    "Operating Cash flow ratio is extremely low compared to its peers",
                "Operating Cash flow ratio is extremely high compared to its peers"};

                points.Add(PointCalculator(AllCompaniesAverage, AllCompaniesSd, companyTick, 1, commentsRatio, "Operating Cash flow ratio"));

                RatioConst(worksheet, "Operating Cash flow ratio", "='CFS'!", "11 /'BS'!", "32", null,  null, col, row, numberOfYears, AllCompanyRatios, AllCompaniesAverage, companyTick, numberOfYearsIncomeStatement,0);
                row += AllCompanyRatios.Count() + 4;
            }
            catch (Exception)
            {

                throw;
            }

            ////Summary liquidity
            //SummaryRatios summary = new SummaryRatios();
            //summary.summary(worksheet, points, comments);

        }
    }
}
