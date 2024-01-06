using Mayntech___Individual_Solution.Auxiliar.Auxiliar.FMP;
using Mayntech___Individual_Solution.Pages;
using OfficeOpenXml.Style;
using OfficeOpenXml;
using System.Drawing;
using Mayntech___Individual_Solution.Auxiliar.Auxiliar.Other;
using OfficeOpenXml.DataValidation;

namespace Mayntech___Individual_Solution.Auxiliar_Analysis.CompetitorsRatios
{
    public class Activity : Ratios
    {
        public void ActivityConstruction(ExcelPackage package, int numberOfYears, List<string> comments,
            string companyTick, int numberOfYearsIncomeStatement)
        {


            List<double> points = new List<double>();

            var worksheet = package.Workbook.Worksheets.Add("Activity");
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

            Dictionary<string, List<double>> DPO = new Dictionary<string, List<double>>();
            Dictionary<string, List<double>> DIO = new Dictionary<string, List<double>>();
            Dictionary<string, List<double>> DSO = new Dictionary<string, List<double>>();

            int firstRatio = 0;
            int SecondRatio = 0;
            int ThirdRatio = 0;
            int numberOfcolumns = Math.Min(5, numberOfYears);

            //DPO
            try
            {
                List<double> ReferenceCompany = new List<double>();
                IDictionary<string, List<double>> AllCompanyRatios = new Dictionary<string, List<double>>();
                IDictionary<string, double> AllCompaniesAverage = new Dictionary<string, double>();
                IDictionary<string, double> AllCompaniesSd = new Dictionary<string, double>();
                //Adiciona os dados da empresa referencia
                for (int i = 0; i < numberOfYears; i++)
                {
                    try
                    {
                        double CompanyAnalysedRatio = (SolutionModel.balances[(numberOfcolumns - 1) - i].AccountPayables / SolutionModel.incomeStatement[(numberOfcolumns - 1) - i].CostOfRevenue)*365;
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
                        bool test = item.Value.Count() >= Math.Min(5, numberOfYears);
                        if (item.Value.Count() >= Math.Min(5, numberOfYears))
                        {

                            try
                            {
                                for (int i = 0; i < numberOfcolumns; i++)
                                {
                                    double ratio = (item.Value[i].AccountPayables/ item.Value[i].CostOfRevenue) * 365;
                                    individualCompanyRatios.Add(ratio);

                                    
                                }
                                firstRatio += 1;
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
                DPO = (Dictionary<string, List<double>>)AllCompanyRatios;

                //Primeiro comentário é para os casos opostos do suposto, o segundo é para outlier negativo e o terceiro
                //(cont) é para outlier positivo
                List<string> commentsRatio = new List<string> {"Current Ratio is below the average of its competitors",
                    "Current ratio is extremely low compared to its peers",
                "Current Ratio is extremely high compared to its peers"};

                //points.Add(PointCalculator(AllCompaniesAverage, AllCompaniesSd, companyTick, 1, commentsRatio, "Days of Payables Outstanding (DPO)"));

                RatioConst(worksheet, "Days of Payables Outstanding (DPO)", "=('BS'!", "27 /- 'P&L'!", "6)*365", null, null, col, row, numberOfYears, AllCompanyRatios, AllCompaniesAverage, companyTick, numberOfYearsIncomeStatement,1);
                row += AllCompanyRatios.Count() + 4;
            }
            catch (Exception)
            {

                throw;
            }


            //DIO
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
                        double CompanyAnalysedRatio = (SolutionModel.balances[(numberOfcolumns - 1) - i].Inventory / SolutionModel.incomeStatement[(numberOfcolumns - 1) - i].CostOfRevenue) * 365;
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
                                    double ratio = (item.Value[i].Inventory / item.Value[i].CostOfRevenue) * 365;
                                    individualCompanyRatios.Add(ratio);

                                    
                                }
                                SecondRatio += 1;
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
                DIO = (Dictionary<string, List<double>>)AllCompanyRatios;

                //Primeiro comentário é para os casos opostos do suposto, o segundo é para outlier negativo e o terceiro
                //(cont) é para outlier positivo
                List<string> commentsRatio = new List<string> {"DPO is below the average of its competitors",
                    "DPO is extremely low compared to its peers",
                "DPO is extremely high compared to its peers"};

                //points.Add(PointCalculator(AllCompaniesAverage, AllCompaniesSd, companyTick, 1, commentsRatio, "Days of Inventory Outstanding (DIO)"));

                RatioConst(worksheet, "Days of Inventory Outstanding (DIO)", "=('BS'!", "10 /- 'P&L'!", "6)*365", null, null, col, row, numberOfYears, AllCompanyRatios, AllCompaniesAverage, companyTick, numberOfYearsIncomeStatement, 1);
                row += AllCompanyRatios.Count() + 4;
            }
            catch (Exception)
            {

                throw;
            }



            //DSO
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
                        double CompanyAnalysedRatio = (SolutionModel.balances[(numberOfcolumns - 1) - i].NetReceivables / SolutionModel.incomeStatement[(numberOfcolumns - 1) - i].CostOfRevenue) * 365;
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
                                    double ratio = (item.Value[i].NetReceivables / item.Value[i].CostOfRevenue) * 365;
                                    individualCompanyRatios.Add(ratio);

                                    
                                }
                                ThirdRatio += 1;
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
                DSO = (Dictionary<string, List<double>>)AllCompanyRatios;

                //Primeiro comentário é para os casos opostos do suposto, o segundo é para outlier negativo e o terceiro
                //(cont) é para outlier positivo
                List<string> commentsRatio = new List<string> {"DSO is below the average of its competitors",
                    "DSO is extremely low compared to its peers",
                "DSO is extremely high compared to its peers"};

                //points.Add(PointCalculator(AllCompaniesAverage, AllCompaniesSd, companyTick, 1, commentsRatio, "Days of Sales Outstanding (DSO)"));

                RatioConst(worksheet, "Days of Sales Outstanding (DSO)", "=('BS'!", "9 / -'P&L'!", "6)*365", null, null, col, row, numberOfYears, AllCompanyRatios, AllCompaniesAverage, companyTick, numberOfYearsIncomeStatement, 1);
                row += AllCompanyRatios.Count() + 4;
            }
            catch (Exception)
            {

                throw;
            }





            //CCC (cash conversion cycle)

            

            Dictionary<string, List<double>> ccc = CashConversionCycle(DPO, DIO, DSO);
            
            for (int i = 0; i < numberOfcolumns + 3; i++)
            {
                worksheet.Cells[row, col + i].Style.Fill.PatternType = ExcelFillStyle.Solid;
                worksheet.Cells[row, col + i].Style.Fill.BackgroundColor.SetColor(Cores.CorPrincipal);

                if (i==0)
                {
                    worksheet.Cells[row, col].Value = "Cash Conversion Cycle";
                    worksheet.Cells[row, col].Style.Font.Bold = true;
                    worksheet.Cells[row, col].Style.Font.Color.SetColor(Color.White);

                    // cria o header
                    worksheet.Cells[row + 1, col].Value = "Ticker";
                    worksheet.Cells[row + 1, col].Style.Font.Bold = true;
                    worksheet.Cells[row + 1, col].Style.Font.Color.SetColor(Cores.CorTexto);
                    worksheet.Cells[row + 1, col].Style.Border.Bottom.Style = ExcelBorderStyle.Thin;

                    //Coluna dos comments
                    worksheet.Cells[row + 1, col + numberOfcolumns + 4].Value = "Comments";
                    worksheet.Cells[row + 1, col + numberOfcolumns + 4].Style.Font.Bold = true;
                    worksheet.Cells[row + 1, col + numberOfcolumns + 4].Style.Border.Bottom.Style = ExcelBorderStyle.Thin;
                    worksheet.Columns[col + numberOfYears + 3 + 4].Width = 50;

                    int a = 0;

                    foreach (KeyValuePair<string, List<double>> item in ccc)
                    {
                        
                        worksheet.Cells[row + 2 + a, col].Value = item.Key;

                        if (item.Key == companyTick)
                        {
                            for (int c = 0; c < numberOfcolumns + 3; c++)
                            {
                                worksheet.Cells[row+2+a, col + c].Style.Fill.PatternType = ExcelFillStyle.Solid;
                                worksheet.Cells[row+2+a, col + c].Style.Fill.BackgroundColor.SetColor(Cores.CorBackgroundAnalysis);
                            }

                        }

                        a += 1;
                    }

                }
                if (i>0 && i <numberOfcolumns +1)
                {
                    ExcelNextCol nextCol = new ExcelNextCol();

                    int aux = numberOfYears - (numberOfcolumns - i) + 2;
                    string column = nextCol.GetExcelColumnName(aux);
                    worksheet.Cells[row + 1, col  + i].Formula = "=YEAR('P&L'!" + column + "3)";
                    worksheet.Cells[row + 1, col+i].Style.Font.Bold = true;
                    worksheet.Cells[row + 1, col+i].Style.Font.Color.SetColor(Cores.CorTexto);
                    worksheet.Cells[row + 1, col+i].Style.Border.Bottom.Style = ExcelBorderStyle.Thin;

                    for (int b = 0; b < ccc.Count(); b++)
                    {
                        int EndFirstRatiotLine = 11 + firstRatio;

                        int BeggSecondRatio = EndFirstRatiotLine + 5;
                        int EndSecondRatio = BeggSecondRatio + SecondRatio;

                        int BeggThirdRatio = EndSecondRatio + 5;
                        int EndThirdRatio = BeggThirdRatio + ThirdRatio;

                        int rowAux = row + 2 + b;
                        int yearColumn = 1 + i;

                        worksheet.Cells[row + 2 + b, col+i].Formula = "=-VLOOKUP(B" + rowAux + ",B11:I" + EndFirstRatiotLine + "," + yearColumn + ",0) + " +
                            "VLOOKUP(B" + rowAux + ",B" + BeggSecondRatio+":I" + EndSecondRatio + "," + yearColumn + ",0)"+ 
                            "+VLOOKUP(B" + rowAux + ",B" + BeggThirdRatio + ":I" + EndThirdRatio + "," + yearColumn + ",0)";

                        worksheet.Cells[row + 2 + b, col + i].Style.Numberformat.Format = "#,##0;(#,##0);-";

                        if (i == numberOfcolumns)
                        {
                            string ColFormula = nextCol.GetExcelColumnName(numberOfcolumns + 2);
                            int rowFormula = row + 2 + b;
                            string error = '"' + "n.a." + '"';
                            //Faz aqui a averga e o coefficient of variation
                            worksheet.Cells[row + 2 + b, col+i+1].Formula = "=IFERROR(AVERAGE(C" + rowFormula + ":" + ColFormula + rowFormula + ")," + error + ")";
                            worksheet.Cells[row + 2 + b, col + i + 2].Formula = "=IFERROR(STDEV(C" + rowFormula + ":" + ColFormula + rowFormula + ")/AVERAGE(C" + rowFormula + ":" + ColFormula + rowFormula + ")," + error + ")";

                            worksheet.Cells[row + 2 + b, col + i + 1].Style.Numberformat.Format = "#,##0;(#,##0);-";
                            worksheet.Cells[row + 2+b, col + i + 1].Style.HorizontalAlignment = ExcelHorizontalAlignment.Right;
                            worksheet.Cells[row + 2+b, col + i + 2].Style.HorizontalAlignment = ExcelHorizontalAlignment.Right;
                            worksheet.Cells[row + 2 + b, col + i + 2].Style.Numberformat.Format = "0.00%";
                        }

                    }
                }
                if (i >=numberOfcolumns+1)
                {
                    worksheet.Cells[row + 1, col + numberOfcolumns + 1].Value = "Average";
                    worksheet.Cells[row + 1, col + numberOfcolumns + 2].Value = "Coefficient of Variation";
                    worksheet.Cells[row + 1, col + numberOfcolumns + 1].Style.Font.Color.SetColor(Cores.CorTexto);
                    worksheet.Cells[row + 1, col + numberOfcolumns + 1].Style.Font.Bold = true;

                    worksheet.Cells[row + 1, col + numberOfcolumns + 2].Style.Font.Color.SetColor(Cores.CorTexto);
                    worksheet.Cells[row + 1, col + numberOfcolumns + 2].Style.Font.Bold = true;

                    worksheet.Cells[row + 1, col + numberOfcolumns + 1].Style.Border.Bottom.Style = ExcelBorderStyle.Thin;
                    worksheet.Cells[row + 1, col + numberOfcolumns + 2].Style.Border.Bottom.Style = ExcelBorderStyle.Thin;
                    worksheet.Columns[col + numberOfcolumns + +2].Width = 25;
                    worksheet.Cells[row + 1, col + numberOfcolumns + 2].Style.HorizontalAlignment = ExcelHorizontalAlignment.Right;
                    worksheet.Columns[col + numberOfcolumns + 1].Width = 12;
                    worksheet.Cells[row + 1, col + numberOfcolumns + 1].Style.HorizontalAlignment = ExcelHorizontalAlignment.Right;

                    

                }
            }

            ////Summary Activity
            //SummaryRatios summary = new SummaryRatios();
            //summary.summary(worksheet, points, comments);

        }

        public Dictionary<string, List<double>> CashConversionCycle(Dictionary<string, List<double>> DPO, Dictionary<string, List<double>> DIO, Dictionary<string, List<double>> DSO)
        {
            Dictionary<string, List<double>> outputAux = new Dictionary<string, List<double>>();
            Dictionary<string, List<double>> output = new Dictionary<string, List<double>>();

            Dictionary<string, double> averageAux = new Dictionary<string, double>();
            foreach (KeyValuePair<string, List<double>> item in DPO)
            {
                try
                {
                    List<double> ccc = new List<double>();
                    for (int i = 0; i < item.Value.Count(); i++)
                    {
                        ccc.Add(-item.Value[i] + DIO[item.Key][i] + DSO[item.Key][i]);
                    }

                    ccc.Reverse();

                    averageAux.Add(item.Key, ccc.Average());

                    outputAux.Add(item.Key, ccc);
                }
                catch 
                {

                    
                }

                
            }
            foreach (var item in averageAux.OrderByDescending(key => key.Value))
            {
                output.Add(item.Key, outputAux[item.Key]);
            }



            return output;   
            
        }
    }
    
}
