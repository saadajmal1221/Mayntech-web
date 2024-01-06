using Mayntech___Individual_Solution.Auxiliar.Auxiliar.FMP;
using Mayntech___Individual_Solution.Pages;
using OfficeOpenXml.Style;
using OfficeOpenXml;
using System.Drawing;

namespace Mayntech___Individual_Solution.Auxiliar_Analysis.CompetitorsRatios
{
    public class Profitability : Ratios
    {

        public void ProfitabilityConstruction(ExcelPackage package, int numberOfYears, List<string> comments,
            string companyTick, int numberOfYearsIncomeStatement)
        {


            List<double> points = new List<double>();

            var worksheet = package.Workbook.Worksheets.Add("Profitability");
            worksheet.View.ShowGridLines = false;
            worksheet.View.ZoomScale = 80;

            for (int i = 0; i < 5; i++)
            {
                worksheet.Rows[2 + i].OutlineLevel = 1;
                worksheet.Rows[2 + i].Collapsed = true;
            }


            int row = 9;
            int col = 2;

            worksheet.Row(7).Style.Fill.PatternType = ExcelFillStyle.Solid;
            worksheet.Row(7).Style.Fill.BackgroundColor.SetColor(Color.Gray);
            worksheet.Cells[7, 2].Value = "Ratios";
            worksheet.Cells[7, 2].Style.Font.Color.SetColor(Color.White);
            worksheet.Cells[7, 2].Style.Font.Bold = true;
            int numberOfcolumns = Math.Min(5, numberOfYears);

            //Operating ROA
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
                        double CompanyAnalysedRatio = SolutionModel.incomeStatement[(numberOfcolumns - 1) - i].OperatingIncome / SolutionModel.balances[(numberOfcolumns - 1) - i].TotalAssets;
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
                                    double ratio = item.Value[i].OperatingIncome / item.Value[i].TotalAssets;
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
                List<string> commentsRatio = new List<string> {"Operating ROA is below the average of its competitors",
                    "Operating ROA is extremely low compared to its peers",
                "Operating ROA is extremely high compared to its peers"};

                points.Add(PointCalculator(AllCompaniesAverage, AllCompaniesSd, companyTick, 1, commentsRatio, "Operating ROA"));

                RatioConst(worksheet, "Operating ROA", "='P&L'!", "15 / 'BS'!", "23", null, null, col, row, numberOfYears, AllCompanyRatios, AllCompaniesAverage, companyTick, numberOfYearsIncomeStatement,0);
                row += AllCompanyRatios.Count() + 4;
            }
            catch (Exception)
            {

                throw;
            }




            ////return on total Capital
            //try
            //{
            //    List<double> ReferenceCompany = new List<double>();
            //    IDictionary<string, List<double>> AllCompanyRatios = new Dictionary<string, List<double>>();
            //    IDictionary<string, double> AllCompaniesAverage = new Dictionary<string, double>();
            //    IDictionary<string, double> AllCompaniesSd = new Dictionary<string, double>();
            //    //Adiciona os dados da empresa referencia
            //    for (int i = 0; i < 5; i++)
            //    {
            //        try
            //        {
            //            double CompanyAnalysedRatio = ((double)SolutionModel.incomeStatement[4 - i].EBITDA - (double)SolutionModel.incomeStatement[4 - i].DepreciationAndAmortization) / (SolutionModel.balances[4 - i].TotalDebt + SolutionModel.balances[4 - i].TotalEquity);
            //            ReferenceCompany.Add(CompanyAnalysedRatio);
            //        }
            //        catch (Exception)
            //        {

            //        }

            //    }

            //    AllCompanyRatios.Add(companyTick, ReferenceCompany);
            //    AllCompaniesAverage.Add(companyTick, ReferenceCompany.Average());
            //    AllCompaniesSd.Add(companyTick, StandardDeviation(ReferenceCompany));




            //    //Adiciona os dados de todas as empresas
            //    foreach (KeyValuePair<string, List<FinancialStatements>> item in SolutionModel.BalanceSheetDict)
            //    {
            //        List<double> individualCompanyRatios = new List<double>();

            //        try
            //        {
            //            if (item.Value.Count() >= Math.Min(5, numberOfYears) && item.Value[0].totalCurrentLiabilities > 0)
            //            {

            //                try
            //                {
            //                    for (int i = 0; i < 5; i++)
            //                    {
            //                        double ratio = ((double)item.Value[i].EBITDA - (double)item.Value[i].DepreciationAndAmortization) / (item.Value[i].TotalEquity + item.Value[i].TotalDebt);
            //                        individualCompanyRatios.Add(ratio);
            //                    }
            //                }
            //                catch (Exception)
            //                {

            //                    continue;
            //                }

            //                //Adiciona os ratios, média e Sd da empresa à lista
            //                AllCompanyRatios.Add(item.Key, individualCompanyRatios);
            //                AllCompaniesAverage.Add(item.Key, individualCompanyRatios.Average());
            //                AllCompaniesSd.Add(item.Key, StandardDeviation(individualCompanyRatios));
            //            }
            //        }
            //        catch (Exception)
            //        {

            //            continue;
            //        }


            //    }

            //    //Primeiro comentário é para os casos opostos do suposto, o segundo é para outlier negativo e o terceiro
            //    //(cont) é para outlier positivo
            //    List<string> commentsRatio = new List<string> {"Return on total capital is below the average of its competitors",
            //        "Return on total capital is extremely low compared to its peers",
            //    "Return on total capital is extremely high compared to its peers"};

            //    points.Add(PointCalculator(AllCompaniesAverage, AllCompaniesSd, companyTick, 1, commentsRatio, "Return on total capital"));

            //    RatioConst(worksheet, "Return on total capital", "=('P&L'!", "27 + 'P&L'!", "29)/('BS'!", "58 + 'BS'!", "53)", col, row, numberOfYears, AllCompanyRatios, AllCompaniesAverage, companyTick, numberOfYearsIncomeStatement, 0);
            //    row += AllCompanyRatios.Count() + 4;
            //}
            //catch (Exception)
            //{

            //    throw;
            //}

            ////Summary Profitability
            //SummaryRatios summary = new SummaryRatios();
            //summary.summary(worksheet, points, comments);

        }




    }
}

