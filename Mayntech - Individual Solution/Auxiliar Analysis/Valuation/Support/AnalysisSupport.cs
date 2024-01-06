using Mayntech___Individual_Solution.Auxiliar.Analysis;
using Mayntech___Individual_Solution.Auxiliar.Auxiliar.Other;
using OfficeOpenXml;
using OfficeOpenXml.Style;

namespace Mayntech___Individual_Solution.Auxiliar.Valuation.Support
{
    public class AnalysisSupport
    {
        public void Estimations(ExcelWorksheet worksheet, int row, int col, List<double> revenue, List<double> value)
        {
            double size = value.Average();
            double slope = 0;

            List<int> confidence = new List<int>();
            double Rsquared = 0;

            double estimate = 0;

            List<double> FiveYears = new List<double>();
            List<double> ThreeYears = new List<double>();
            List<double> TotalXvalues = new List<double>();
            List<double> ThreeXvalues = new List<double>();
            List<double> FiveXvalues = new List<double>();

            for (int i = 0; i < revenue.Count(); i++)
            {
                TotalXvalues.Add(i);

                if (revenue.Count()>5)
                {
                    if (revenue.Count()-1-i <5)
                    {
                        FiveYears.Add(value[i]);
                        FiveXvalues.Add(i-3);
                    }
                }
                if (revenue.Count() > 3)
                {
                    if (revenue.Count() - 1 - i < 3)
                    {
                        ThreeYears.Add(value[i]);
                        ThreeXvalues.Add(i-5);
                    }
                }
            }


            LinearRegression linearRegression = new LinearRegression();
            List<double> LrThreeYears = linearRegression.LinearRegressionCalculation(ThreeXvalues, ThreeYears);
            List<double> LrFiveYears = linearRegression.LinearRegressionCalculation(FiveXvalues, FiveYears);
            List<double> LrTotal = linearRegression.LinearRegressionCalculation(TotalXvalues, value);

            List<string> order = new List<string>();
            if (ThreeYears.Count()>0 && FiveYears.Count()>0) //Se conseguir fazer as 2 listas
            {
                if (LrFiveYears[2] > LrTotal[2])
                {
                    if (LrFiveYears[2]> LrThreeYears[2])
                    {
                        if (LrThreeYears[2]> LrTotal[2])
                        {
                            //Five, Three, total
                            order.AddRange(new List<string> { "Five", "Three", "Total" });
                        }
                        else
                        {
                            //Five, total, Three
                            order.AddRange(new List<string> { "Five", "Total", "Three"});
                        }
                    }
                    else
                    {
                        //Three, five, total
                        order.AddRange(new List<string> { "Three", "Five", "Total"  });
                    }

                }
                else
                {
                    if (LrTotal[2] > LrThreeYears[2])
                    {
                        if (LrThreeYears[2] > LrFiveYears[2])
                        {
                            //total, Three, five
                            order.AddRange(new List<string> { "Total", "Three", "Five" });
                        }
                        else
                        {
                            //total, five, Three
                            order.AddRange(new List<string> { "Total", "Five","Three" });
                        }
                    }
                    else
                    {
                        //Three, total, five
                        order.AddRange(new List<string> {"Three","Total", "Five" });
                    }
                }
            }
            else if (ThreeYears.Count()>0) //Se não tiver os dados para a lista dos 5 anos
            {
                confidence.Add(2);

                if (LrThreeYears[2] > LrTotal[2])
                {
                    order.AddRange(new List<string> { "Three", "Total" });
                }
                else
                {
                    order.AddRange(new List<string> {  "Total", "Three"});
                }
            }
            else //Se não tiver dados para fazer nenhuma lista
            {
                confidence.Add(3);
                order.Add("Total");
            }
            slope = 0.1;
            Rsquared = 0.9;

            //else
            //{
            //    slope = 0.3;
            //    Rsquared = 0.6;
            //}

            //Vê qual é o número a utilizar
            double rsquaredAux = 0;
            if (order[0] == "Total")
            {
                rsquaredAux = LrTotal[2];
            }
            else if (order[0] == "Five")
            {
                rsquaredAux = LrFiveYears[2];
            }
            else
            {
                rsquaredAux = LrThreeYears[2];
            }



            //Adiciona a confiança tendo em conta o rsquared
            if (rsquaredAux>Rsquared)
            {
                confidence.Add(1);
            }
            else if (rsquaredAux > Rsquared/2)
            {
                confidence.Add(2);
            }
            else
            {
                confidence.Add(3);
            }


            //Define o valor estimado
            if (order[0] == "Three")
            {
                estimate = (LrThreeYears[0] * 3 + LrThreeYears[1])*0.5 + value.Average()*0.5;
            }
            else if (order[0] == "Five")
            {
                double aux = (LrFiveYears[0] * 5 + LrFiveYears[1]) * 0.5 + FiveYears.Average()*0.5;
                estimate = aux;
            }
            else
            {
                double aux = (LrTotal[0] * (revenue.Count()+1) + LrTotal[1]) * 0.5 + value.Average()*0.5;
                estimate = aux;
            }


            CommentAndColorValuation(worksheet, confidence, estimate, row, col);

        }
        public void CommentAndColorValuation(ExcelWorksheet worksheet, List<int> confidence, double estimate, int row, int column)
        {
            int confidenceLevel = confidence.Max();
            double min = 0;
            double max = 0;

            if (confidenceLevel ==3)
            {
                min = estimate - 0.03;
                max = estimate + 0.03;

                worksheet.Cells[row+1, column].Style.Fill.PatternType = ExcelFillStyle.Solid; 
                worksheet.Cells[row+1, column].Style.Fill.BackgroundColor.SetColor(Cores.OrangeWarning);

                worksheet.Cells[row + 1, column + 1].Value = "[" + ((short)(min * 100)) + "% ; " + ((short)(max * 100)) + "%]";
                worksheet.Cells[row + 1, column + 2].Value = estimate;
                worksheet.Cells[row + 1, column + 2].Style.Numberformat.Format = "0.00%";
                worksheet.Cells[row + 1, column + 2].Style.HorizontalAlignment = ExcelHorizontalAlignment.Left;


            }
            else if (confidenceLevel == 2)
            {
                min = estimate - 0.02;
                max = estimate + 0.02;

                worksheet.Cells[row + 1, column].Style.Fill.PatternType = ExcelFillStyle.Solid;
                worksheet.Cells[row + 1, column].Style.Fill.BackgroundColor.SetColor(Cores.YellowWarning);

                worksheet.Cells[row + 1, column + 1].Value = "[" + ((short)(min * 100)) + "% ; " + ((short)(max * 100)) + "%]";
                worksheet.Cells[row + 1, column + 2].Value =estimate;
                worksheet.Cells[row + 1, column + 2].Style.Numberformat.Format = "0.00%";
                worksheet.Cells[row + 1, column + 2].Style.HorizontalAlignment = ExcelHorizontalAlignment.Left;

            }
            else
            {
                min = estimate - 0.02;
                max = estimate + 0.02;

                worksheet.Cells[row + 1, column].Style.Fill.PatternType = ExcelFillStyle.Solid;
                worksheet.Cells[row + 1, column].Style.Fill.BackgroundColor.SetColor(Cores.LightGreenWarning);

                worksheet.Cells[row + 1, column + 1].Value = "[" + ((short)(min * 100)) + "% ; " + ((short)(max * 100)) + "%]";
                worksheet.Cells[row + 1, column + 2].Value = estimate;
                worksheet.Cells[row + 1, column + 2].Style.Numberformat.Format = "0.00%";
                worksheet.Cells[row + 1, column + 2].Style.HorizontalAlignment = ExcelHorizontalAlignment.Left;
            }
        }
    }
}
