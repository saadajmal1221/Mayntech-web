using Mayntech___Individual_Solution.Auxiliar.Auxiliar.FMP;
using Mayntech___Individual_Solution.Pages;
using OfficeOpenXml.Style;
using OfficeOpenXml;
using System.Drawing;

namespace Mayntech___Individual_Solution.Auxiliar_Analysis.CompetitorsRatios
{
    public class Solvency : Ratios
    {
        public void SolvencyConstruction(ExcelPackage package, int numberOfYears, List<string> comments,
            string companyTick, int numberOfYearsIncomeStatement)
        {


            List<double> points = new List<double>();

            var worksheet = package.Workbook.Worksheets.Add("Solvency");
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
                        double CompanyAnalysedRatio = SolutionModel.balances[(numberOfcolumns-1) - i].TotalDebt / SolutionModel.balances[(numberOfcolumns - 1) - i].TotalEquity;
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
                        if (item.Value.Count() >= Math.Min(5, numberOfYears) && item.Value[0].TotalEquity > 0)
                        {

                            try
                            {
                                for (int i = 0; i < numberOfcolumns; i++)
                                {
                                    double ratio = item.Value[i].TotalDebt / item.Value[i].TotalEquity;
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

                string ratioName = "Debt to Equity";
                //Primeiro comentário é para os casos opostos do suposto, o segundo é para outlier negativo e o terceiro
                //(cont) é para outlier positivo
                List<string> commentsRatio = new List<string> { ratioName + " is below the average of its competitors",
                    ratioName +" is extremely low compared to its peers",
                ratioName + " is extremely high compared to its peers"};

                points.Add(PointCalculator(AllCompaniesAverage, AllCompaniesSd, companyTick, 2, commentsRatio, ratioName));

                RatioConst(worksheet, ratioName, "='BS'!", "58 / 'BS'!", "53", null, null, col, row, numberOfYears, AllCompanyRatios, AllCompaniesAverage, companyTick, numberOfYearsIncomeStatement,0);
                row += AllCompanyRatios.Count() + 4;
            }
            catch (Exception)
            {

                throw;
            }




            //Debt Ratio
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
                        double CompanyAnalysedRatio = SolutionModel.balances[(numberOfcolumns - 1) - i].TotalDebt / SolutionModel.balances[(numberOfcolumns - 1) - i].TotalAssets;
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
                        if (item.Value.Count() >= Math.Min(5, numberOfYears) && item.Value[0].TotalEquity > 0)
                        {

                            try
                            {
                                for (int i = 0; i < numberOfcolumns; i++)
                                {
                                    double ratio = item.Value[i].TotalDebt / item.Value[i].TotalAssets;
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

                string ratioName = "Debt ratio";
                //Primeiro comentário é para os casos opostos do suposto, o segundo é para outlier negativo e o terceiro
                //(cont) é para outlier positivo
                List<string> commentsRatio = new List<string> { ratioName + " is below the average of its competitors",
                    ratioName +" is extremely low compared to its peers",
                ratioName + " is extremely high compared to its peers"};

                points.Add(PointCalculator(AllCompaniesAverage, AllCompaniesSd, companyTick, 2, commentsRatio, ratioName));

                RatioConst(worksheet, ratioName, "='BS'!", "58 / 'BS'!", "23", null, null, col, row, numberOfYears, AllCompanyRatios, AllCompaniesAverage, companyTick, numberOfYearsIncomeStatement,0);
                row += AllCompanyRatios.Count() + 4;
            }
            catch (Exception)
            {

                throw;
            }

            ////Summary Solvency
            //SummaryRatios summary = new SummaryRatios();
            //summary.summary(worksheet, points, comments);
        }
    }
}
