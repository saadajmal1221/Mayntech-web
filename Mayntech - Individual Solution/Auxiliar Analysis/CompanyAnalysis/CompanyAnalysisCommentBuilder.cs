using Mayntech___Individual_Solution.Auxiliar.Analysis;
using Mayntech___Individual_Solution.Auxiliar.Auxiliar.Other;
using Mayntech___Individual_Solution.Auxiliar.Executive_Summary;
using Mayntech___Individual_Solution.Auxiliar_Analysis.Valuation;
using OfficeOpenXml;
using OfficeOpenXml.Style;
using System.Collections.Generic;

namespace Mayntech___Individual_Solution.Auxiliar.CompanyAnalysis
{
    public class CompanyAnalysisCommentBuilder : CommentAndColor
    {
        public void CommentBuilderIncomeStatement(ExcelWorksheet worksheet, int row, int column, List<double> revenue ,List<double> GrossProfit, 
            List<double> GrossMargin, List<double> OperatingMargin, List<double> OperatingProfit, List<double> Noplat)
        {
            List<string> PositiveComment = new List<string>();
            List<string> NegativeComment = new List<string>();
            List<string> OtherComment = new List<string>();
            List<int> color = new List<int>();

            ValuationSupport support = new ValuationSupport();

            List<double> xvalues = new List<double>();
            for (int i = 0; i < revenue.Count(); i++)
            {
                xvalues.Add(i);
            }         
            LinearRegression linearRegression = new LinearRegression();

            List<double> LRRevenue = linearRegression.LinearRegressionCalculation(xvalues, revenue);
            List<double> LROperatingMargin = linearRegression.LinearRegressionCalculation(xvalues, OperatingMargin);
            List<double> LROperatingProfit = linearRegression.LinearRegressionCalculation(xvalues, OperatingProfit);
            List<double> LrNoplat = linearRegression.LinearRegressionCalculation(xvalues, Noplat);

            string grossMarginAux = GrossMarginEvol(GrossProfit, GrossMargin, xvalues);
            double cagr = support.Cagr(revenue);
            double cgarNoplat = support.Cagr(Noplat);

            if (LRRevenue[0] / revenue.Average() > 0.3)  //Grande aumento nas vendas
            {
                if (LRRevenue[2]>0.3 && cagr >0)
                {
                    PositiveComment.Add("Revenues are rising at a fast pace. ");
                    color.Add(1);

                    AddToExecSummary("Revenues", "DarkGreen", "Revenues are rising at a fast pace. ");
                }
                else
                {
                    PositiveComment.Add("Revenues are rising in a unstable manner. ");
                    color.Add(2);
                    AddToExecSummary("Revenues", "LightGreen", "Revenues are rising in a unstable manner. ");
                }


                if (grossMarginAux == "Fast Increasing Rate")
                {
                    PositiveComment.Add("Gross margin is rising at a fast pace. ");
                    color.Add(1);

                    AddToExecSummary("GrossMargin", "DarkGreen", "Gross margin is rising at a fast pace. ");

                    if (IsDecreasing(LROperatingMargin) == "Yes")
                    {
                        NegativeComment.Add("The company lost operating margin. ");
                        color.Add(3);

                        AddToExecSummary("OperatingMargin", "Yellow", "The company lost operating margin. ");
                    }
                    else if (IsDecreasing(LROperatingProfit) == "Yes")
                    {
                        NegativeComment.Add("The company lost operating profit. ");
                        color.Add(5);

                        AddToExecSummary("OperatingProfit", "Red", "The company lost operating profit. ");
                    }
                }
                else if (grossMarginAux == "Increasing Rate")
                {
                    PositiveComment.Add("The company gained gross margin. ");
                    color.Add(2);

                    AddToExecSummary("GrossMargin", "LightGreen", "The company gained gross margin. ");

                    if (IsDecreasing(LROperatingMargin) == "Yes")
                    {
                        NegativeComment.Add("The company lost operating margin. ");
                        color.Add(3);

                        AddToExecSummary("OperatingMargin", "Yellow", "The company lost operating margin. ");
                    }
                    else if (IsDecreasing(LROperatingProfit) == "Yes")
                    {
                        NegativeComment.Add("The company lost operating profit. ");
                        color.Add(4);

                        AddToExecSummary("OperatingProfit", "Orange", "The company lost operating profit. ");
                    }
                }
                else if (grossMarginAux == "Same Rate")
                {
                    if (IsDecreasing(LROperatingMargin) == "Yes")
                    {
                        NegativeComment.Add("The company lost operating margin. ");
                        color.Add(3);

                        AddToExecSummary("OperatingMargin", "Yellow", "The company lost operating margin. ");
                    }
                }
                else if (grossMarginAux == "Decreasing rate")
                {
                    NegativeComment.Add("The company lost gross margin. ");
                    color.Add(3);

                    AddToExecSummary("GrossMargin", "Yellow", "The company lost gross margin. ");

                    if (IsGrowing(LROperatingMargin) == "Yes")
                    {
                        PositiveComment.Add("The company gained operating margin. ");
                        color.Add(2);

                        AddToExecSummary("OperatingMargin", "LightGreen", "The company gained operating margin. ");
                    }
                }
                else if (grossMarginAux == "Decreasing Absolute")
                {
                    NegativeComment.Add("The rise in revenues is damaging the gross profit. ");
                    color.Add(5);

                    AddToExecSummary("GrossProfit", "Red", "The rise in revenues is damaging the gross profit. ");

                    if (IsGrowing(LROperatingMargin) == "Yes")
                    {
                        PositiveComment.Add("The company gained operating margin. ");
                        color.Add(2);

                        AddToExecSummary("OperatingMargin", "LightGreen", "The company gained operating margin. ");
                    }
                    else if (IsGrowing(LROperatingProfit) == "Yes")
                    {
                        NegativeComment.Add("The company gained operating profit. ");
                        color.Add(2);

                        AddToExecSummary("OperatingProfit", "LightGreen", "The company gained operating profit. ");
                    }
                }
                else
                {
                    NegativeComment.Add("The gross margin of the company is volatile. ");
                    color.Add(3);

                    AddToExecSummary("GrossMargin", "Yellow", "The gross margin of the company is volatile. ");
                }
            }
            else if (LRRevenue[0] / revenue.Average() > 0.03 && cagr >0)  // aumento nas vendas
            {

                if (LRRevenue[2] > 0.3)
                {
                    PositiveComment.Add("Increase in Revenues. ");
                    color.Add(1);

                    AddToExecSummary("Revenues", "DarkGreen", "Increase in Revenues. ");
                }
                else
                {
                    PositiveComment.Add("Revenues are rising in a unstable manner. ");
                    color.Add(2);
                    AddToExecSummary("Revenues", "LightGreen", "Revenues are rising in a unstable manner. ");
                }



                if (grossMarginAux == "Fast Increasing Rate")
                {
                    PositiveComment.Add("Gross margin increased at fast pace. ");
                    color.Add(1);

                    AddToExecSummary("GrossMargin", "DarkGreen", "Gross margin increased at fast pace. ");

                    if (IsDecreasing(LROperatingMargin) == "Yes")
                    {
                        NegativeComment.Add("The company lost operating margin. ");
                        color.Add(3);

                        AddToExecSummary("OperatingMargin", "Yellow", "The company lost operating margin. ");
                    }
                    else if (IsDecreasing(LROperatingProfit) == "Yes")
                    {
                        NegativeComment.Add("The company lost operating profit. ");
                        color.Add(4);

                        AddToExecSummary("OperatingProfit", "Orange", "The company lost operating profit. ");
                    }
                }
                else if (grossMarginAux == "Increasing Rate")
                {
                    PositiveComment.Add("The company gained gross margin. ");
                    color.Add(2);

                    AddToExecSummary("GrossMargin", "LightGreen", "The company gained gross margin. ");

                    if (IsDecreasing(LROperatingMargin) == "Yes")
                    {
                        NegativeComment.Add("The company lost operating margin. ");
                        color.Add(3);

                        AddToExecSummary("OperatingMargin", "Yellow", "The company lost operating margin. ");
                    }
                    else if (IsDecreasing(LROperatingProfit) == "Yes")
                    {
                        NegativeComment.Add("The company lost operating profit. ");
                        color.Add(4);

                        AddToExecSummary("OperatingProfit", "Orange", "The company lost operating profit. ");
                    }
                }
                else if (grossMarginAux == "Same Rate")
                {
                    if (IsDecreasing(LROperatingMargin) == "Yes")
                    {
                        NegativeComment.Add("The company lost operating margin. ");
                        color.Add(3);

                        AddToExecSummary("OperatingMargin", "Yellow", "The company lost operating margin. ");
                    }

                }
                else if (grossMarginAux == "Decreasing rate")
                {
                    NegativeComment.Add("The company lost gross margin. ");
                    color.Add(3);

                    AddToExecSummary("GrossMargin", "Yellow", "The company lost gross margin. ");

                    if (IsGrowing(LROperatingMargin) == "Yes")
                    {
                        PositiveComment.Add("The company gained operating margin. ");
                        color.Add(2);

                        AddToExecSummary("OperatingMargin", "LightGreen", "The company gained operating margin. ");
                    }
                    else if (IsGrowing(LROperatingProfit) == "Yes")
                    {
                        NegativeComment.Add("The company gained operating profit. ");
                        color.Add(2);

                        AddToExecSummary("OperatingProfit", "LightGreen", "The company gained operating profit. ");
                    }
                }
                else if (grossMarginAux == "Decreasing Absolute")
                {
                    NegativeComment.Add("The rise in revenues is damaging the gross profit. ");
                    color.Add(5);

                    AddToExecSummary("GrossProfit", "Red", "The rise in revenues is damaging the gross profit. ");

                    if (IsGrowing(LROperatingMargin) == "Yes")
                    {
                        PositiveComment.Add("The company gained operating margin. ");
                        color.Add(2);

                        AddToExecSummary("OperatingMargin", "LightGreen", "The company gained operating margin. ");
                    }
                    else if (IsGrowing(LROperatingProfit) == "Yes")
                    {
                        NegativeComment.Add("The company gained operating profit. ");
                        color.Add(2);

                        AddToExecSummary("OperatingProfit", "LightGreen", "The company gained operating profit. ");
                    }
                }
                else
                {
                    NegativeComment.Add("The gross margin of the company is volatile. ");
                    color.Add(3);

                    AddToExecSummary("GrossMargin", "Yellow", "The gross margin of the company is volatile. ");
                }

            }
            else if (LRRevenue[0] / revenue.Average() > -0.02 )  // estagnação nas vendas
            {
                
                if (LRRevenue[2] > 0.3)
                {
                    OtherComment.Add("Revenues are relatively stable. ");
                }
                else
                {
                    NegativeComment.Add("Revenues are volatile. ");
                    color.Add(3);

                    AddToExecSummary("Revenues", "Yellow", "Revenues are volatile. ");
                }

                if (grossMarginAux == "Fast Increasing Rate")
                {
                    PositiveComment.Add("Gross margin increased at fast pace. ");
                    color.Add(1);

                    AddToExecSummary("GrossMargin", "DarkGreen", "Gross margin increased at fast pace. ");

                    if (IsDecreasing(LROperatingMargin) == "Yes")
                    {
                        NegativeComment.Add("The company lost operating margin. ");
                        color.Add(3);

                        AddToExecSummary("OperatingMargin", "Yellow", "The company lost operating margin. ");
                    }
                }
                else if (grossMarginAux == "Increasing Rate")
                {
                    PositiveComment.Add("The company gained gross margin. ");
                    color.Add(2);

                    AddToExecSummary("GrossMargin", "LightGreen", "The company gained gross margin. ");

                    if (IsDecreasing(LROperatingMargin) == "Yes")
                    {
                        NegativeComment.Add("The company lost operating margin. ");
                        color.Add(3);

                        AddToExecSummary("OperatingMargin", "Yellow", "The company lost operating margin. ");
                    }

                }
                else if (grossMarginAux == "Decreasing Rate")
                {
                    NegativeComment.Add("The company lost gross margin. ");
                    color.Add(3);

                    AddToExecSummary("GrossMargin", "Yellow", "The company lost gross margin. ");

                    if (IsGrowing(LROperatingMargin) == "Yes")
                    {
                        PositiveComment.Add("The company gained operating margin. ");
                        color.Add(2);

                        AddToExecSummary("OperatingMargin", "LightGreen", "The company gained operating margin. ");
                    }
                }
                else if (grossMarginAux == "Volatile")
                {
                    NegativeComment.Add("The gross margin of the company is volatile. ");
                    color.Add(3);

                    AddToExecSummary("GrossMargin", "Yellow", "The gross margin of the company is volatile. ");

                    if (IsGrowing(LROperatingMargin) == "Yes")
                    {
                        PositiveComment.Add("The company gained operating margin. ");
                        color.Add(2);

                        AddToExecSummary("OperatingMargin", "LightGreen", "The company gained operating margin. ");
                    }
                }
            }
            else if (LRRevenue[0] / revenue.Average() > -0.1 && cagr<0) // decrescimo nas vendas
            {


                if (LRRevenue[2] > 0.3)
                {
                    NegativeComment.Add("Revenues decreased. ");
                    color.Add(5);

                    AddToExecSummary("Revenues", "Red", "Revenues decreased. ");
                }
                else
                {
                    NegativeComment.Add("Revenues decreased in an unstable manner. ");
                    color.Add(5);

                    AddToExecSummary("Revenues", "Red", "Revenues decreased in an unstable manner. ");
                }

                if (grossMarginAux == "Fast Increasing Rate")
                {
                    PositiveComment.Add("Gross margin increased at fast pace. ");
                    color.Add(1);

                    AddToExecSummary("GrossMargin", "DarkGreen", "Gross margin increased at fast pace. ");

                    if (IsDecreasing(LROperatingMargin) == "Yes")
                    {
                        NegativeComment.Add("The company lost operating margin. ");
                        color.Add(3);

                        AddToExecSummary("OperatingMargin", "Yellow", "The company lost operating margin. ");
                    }
                }
                else if (grossMarginAux == "Increasing Rate")
                {
                    PositiveComment.Add("The company gained gross margin. ");
                    color.Add(2);

                    AddToExecSummary("GrossMargin", "LightGreen", "The company gained gross margin. ");

                    if (IsDecreasing(LROperatingMargin) == "Yes")
                    {
                        NegativeComment.Add("The company lost operating margin. ");
                        color.Add(3);

                        AddToExecSummary("OperatingMargin", "Yellow", "The company lost operating margin. ");
                    }

                }
                else if (grossMarginAux == "Decreasing Rate")
                {
                    NegativeComment.Add("The company lost gross margin. ");
                    color.Add(3);

                    AddToExecSummary("GrossMargin", "Yellow", "The company lost gross margin. ");

                    if (IsGrowing(LROperatingMargin) == "Yes")
                    {
                        PositiveComment.Add("The company gained operating margin. ");
                        color.Add(2);

                        AddToExecSummary("OperatingMargin", "LightGreen", "The company gained operating margin. ");
                    }
                }
                else if (grossMarginAux == "Volatile")
                {
                    NegativeComment.Add("The gross margin of the company is volatile. ");
                    color.Add(3);

                    AddToExecSummary("GrossMargin", "Yellow", "The gross margin of the company is volatile. ");

                    if (IsGrowing(LROperatingMargin) == "Yes")
                    {
                        PositiveComment.Add("The company gained operating margin. ");
                        color.Add(2);

                        AddToExecSummary("OperatingMargin", "LightGreen", "The company gained operating margin. ");
                    }
                }
            }
            else
            {


                if (LRRevenue[2] > 0.3 && cagr < 0)
                {
                    NegativeComment.Add("Revenues decreased at an unsettling pace. ");
                    color.Add(5);

                    AddToExecSummary("Revenues", "Red", "Revenues decreased at an unsettling pace. ");
                }
                else
                {
                    if (cagr < 0)
                    {
                        NegativeComment.Add("Revenues decreased at an unsettling pace and unstable manner. ");
                        color.Add(5);

                        AddToExecSummary("Revenues", "Red", "Revenues decreased at an unsettling pace and unstable manner. ");
                    }

                }

                if (grossMarginAux == "Fast Increasing Rate")
                {
                    PositiveComment.Add("Gross margin increased at fast pace. ");
                    color.Add(1);

                    AddToExecSummary("GrossMargin", "DarkGreen", "Gross margin increased at fast pace. ");

                    if (IsDecreasing(LROperatingMargin) == "Yes")
                    {
                        NegativeComment.Add("The company lost operating margin. ");
                        color.Add(3);

                        AddToExecSummary("OperatingMargin", "Yellow", "The company lost operating margin. ");
                    }
                }
                else if (grossMarginAux == "Increasing Rate")
                {
                    PositiveComment.Add("The company gained gross margin. ");
                    color.Add(2);

                    AddToExecSummary("GrossMargin", "LightGreen", "The company gained gross margin. ");

                    if (IsDecreasing(LROperatingMargin) == "Yes")
                    {
                        NegativeComment.Add("The company lost operating margin. ");
                        color.Add(3);

                        AddToExecSummary("OperatingMargin", "Yellow", "The company lost operating margin. ");
                    }

                }
                else if (grossMarginAux == "Decreasing Rate")
                {
                    NegativeComment.Add("The company lost gross margin. ");
                    color.Add(3);

                    AddToExecSummary("GrossMargin", "Yellow", "The company lost gross margin. ");

                    if (IsGrowing(LROperatingMargin) == "Yes")
                    {
                        PositiveComment.Add("The company gained operating margin. ");
                        color.Add(2);

                        AddToExecSummary("OperatingMargin", "LightGreen", "The company gained operating margin. ");
                    }
                }
                else if (grossMarginAux == "Volatile")
                {
                    NegativeComment.Add("The gross margin of the company is volatile. ");
                    color.Add(3);

                    AddToExecSummary("GrossMargin", "Yellow", "The gross margin of the company is volatile. ");

                    if (IsGrowing(LROperatingMargin) == "Yes")
                    {
                        PositiveComment.Add("The company gained operating margin. ");
                        color.Add(2);

                        AddToExecSummary("OperatingMargin", "LightGreen", "The company gained operating margin. ");
                    }
                }
            }

            List<string> NoplatResult = simpleAnalysis(LrNoplat, cgarNoplat, "NOPLAT");

            if (NoplatResult[1] == "++")
            {
                PositiveComment.Add(NoplatResult[0]);
                color.Add(1);

                AddToExecSummary("NOPLAT", "DarkGreen", "The company's NOPLAT is rising at a fast pace. ");
            }
            else if (NoplatResult[1] == "+")
            {
                PositiveComment.Add(NoplatResult[0]);
                color.Add(2);

                AddToExecSummary("NOPLAT", "LightGreen", "The company's NOPLAT is rising. ");
            }
            else if (NoplatResult[1] == "+-")
            {
                AddToExecSummary("NOPLAT", "Yellow", "The company's NOPLAT is relatively stable. ");
            }
            else if (NoplatResult[1] == "-")
            {
                NegativeComment.Add(NoplatResult[0]);
                color.Add(3);

                AddToExecSummary("NOPLAT", "Orange", "The company's NOPLAT is decreasing. ");
            }
            else if (NoplatResult[1] == "--")
            {
                NegativeComment.Add(NoplatResult[0]);
                color.Add(4);

                AddToExecSummary("NOPLAT", "Red", "The company's NOPLAT is decreasing at a fast pace. ");
            }

            if (NoplatResult[2] == "High volatility")
            {
                PositiveComment.Add("High NOPLAT volatility");
                color.Add(4);

                AddToExecSummary("NOPLAT", "Red", "The company's NOPLAT registered a high volatility. It is important to understand why this is happening.  ");
            }

            Comment(worksheet, 4, column, color, PositiveComment, NegativeComment, OtherComment);
            worksheet.Column(12).Width = 10;
        }


            


        

        public void CommentBuilderBalanceSheet(ExcelWorksheet worksheet, int row, int column, List<int> years, List<double> WorkingCapital, 
            List<double> totalDebt, List<double> Cash, List<double> PPE
            , List<double> Goodwill, List<double> TotalAssets)
        {
            List<string> PositiveComment = new List<string>();
            List<string> NegativeComment = new List<string>();
            List<string> OtherComment = new List<string>();
            List<int> color = new List<int>();

            CalculationsAux values = new CalculationsAux();
            List<int> negativeWorkingCapital = values.FindNegativeValues(years, WorkingCapital);


            StandardDeviation standardDeviation = new StandardDeviation();
            double CoefficientVariationWC = standardDeviation.Sd(WorkingCapital)/WorkingCapital.Average();

            ValuationSupport support = new ValuationSupport();
            double cagrWC = support.Cagr(WorkingCapital);


            //Lista dos valores em que a Debt e o Goodwill ultrapassaram as referencias
            List<int> GoodwillRef = values.PercentageRef(TotalAssets, Goodwill, 0.3);

            List<double> xvalues = new List<double>();
            for (int i = 0; i < WorkingCapital.Count(); i++)
            {
                xvalues.Add(i);
            }
            LinearRegression LR = new LinearRegression();
            List<double> LRWOrkingCapital = LR.LinearRegressionCalculation(xvalues, WorkingCapital);
            List<double> LRTotalDebt = LR.LinearRegressionCalculation(xvalues, totalDebt);
            List<double> LRCash = LR.LinearRegressionCalculation(xvalues, Cash);
            List<double> LRPPE = LR.LinearRegressionCalculation(xvalues, PPE);


            //Cash
            if (LRCash[2]>0.5)
            {
                if (LRCash[0] / Cash.Average() > 0.4)
                {
                    OtherComment.Add("The company accumulated Cash at a fast pace. ");
                    
                }
                else if (LRCash[0]/Cash.Average()>0.1)
                {
                    PositiveComment.Add("The company accumulated Cash. ");
                    color.Add(2);

                    AddToExecSummary("Cash", "LightGreen", "The company accumulated Cash. ");
                }
                else if (LRCash[0] / Cash.Average() < -0.1)
                {
                    NegativeComment.Add("The company lost a significant amount of Cash. ");
                    color.Add(4);

                    AddToExecSummary("Cash", "Orange", "The company lost a significant amount of Cash. ");
                }
                else if (LRCash[0] / Cash.Average() < -0.4)
                {
                    NegativeComment.Add("The company lost a significant amount of Cash. ");
                    color.Add(5);

                    AddToExecSummary("Cash", "Red", "The company lost a significant amount of Cash. ");
                }
            }
            else
            {
                if (LRCash[0] / Cash.Average() > 0.6)
                {
                    OtherComment.Add("The company accumulated Cash at an unsettling pace. ");
                    
                }
                else if (LRCash[0] / Cash.Average() > 0.2)
                {
                    PositiveComment.Add("The company accumulated Cash. ");
                    color.Add(2);

                    AddToExecSummary("Cash", "LightGreen", "The company accumulated Cash. ");
                }
                else if (LRCash[0] / Cash.Average() < -0.2)
                {
                    NegativeComment.Add("The company lost a significant amount of Cash. ");
                    color.Add(4);

                    AddToExecSummary("Cash", "Orange", "The company lost a significant amount of Cash. ");
                }
                else if (LRCash[0] / Cash.Average() < -0.6)
                {
                    NegativeComment.Add("The company lost a significant amount of Cash. ");
                    color.Add(5);

                    AddToExecSummary("Cash", "Red", "The company lost a significant amount of Cash. ");
                }
            }


            //PP&E
            if (LRPPE[2] > 0.6)
            {
                if (LRPPE[0]/PPE.Average()>0.3)
                {
                    OtherComment.Add("The company acquired a significant amount of PP&E. ");
                    color.Add(2);

                    AddToExecSummary("PP&E", "LightGreen", "The company acquired a significant amount of PP&E. Are the company's operations dependent on a large amount of PP&E? Or is this increase in PP&E related with a change in strategy?");
                }
                else if (LRPPE[0] / PPE.Average() > 0.1)
                {
                    OtherComment.Add("The company acquired PP&E. ");
                    color.Add(2);

                    AddToExecSummary("PP&E", "LightGreen", "The company acquired PP&E. Are the company's operations dependent on a large amount of PP&E? ");
                }
                else if (LRPPE[0] / PPE.Average() < -0.2)
                {
                    NegativeComment.Add("The company lost a significant amount of PP&E. ");
                    color.Add(5);

                    AddToExecSummary("PP&E", "Red", "The company lost a significant amount of PP&E. Is the company divesting its physical assets? If yes, then why and which assets are being sold?");
                }
            }
            else
            {
                if (LRPPE[0] / PPE.Average() > 0.5)
                {
                    OtherComment.Add("The company acquired a significant amount of PP&E. ");
                    color.Add(2);

                    AddToExecSummary("PP&E", "LightGreen", "The company acquired a significant amount of PP&E. Are the company's operations dependent on a large amount of PP&E? Or is this increase in PP&E related with a change in strategy? ");
                }
                else if (LRPPE[0] / PPE.Average() < -0.4)
                {
                    NegativeComment.Add("The company lost a significant amount of PP&E. ");
                    color.Add(4);

                    AddToExecSummary("PP&E", "Orange", "The company lost a significant amount of PP&E. Is the company divesting its physical assets? If yes, then why and which assets are being sold?\"");
                }
            }



            //WorkingCapital
            List<string> NoplatResult = simpleAnalysis(LRWOrkingCapital, cagrWC, "Operational Working Capital");

            if (NoplatResult[1] == "++")
            {
                NegativeComment.Add(NoplatResult[0]);
                color.Add(3);

                AddToExecSummary("OpWorkingCapital", "Yellow", "The company's Operational Working Capital is rising at a fast pace. What is the reason for this increase?  ");
            }
            else if (NoplatResult[1] == "+")
            {
                NegativeComment.Add(NoplatResult[0]);
                color.Add(3);

                AddToExecSummary("OpWorkingCapital", "Yellow", "The company's Operational Working Capital is rising. ");
            }
            else if (NoplatResult[1] == "+-")
            {
                AddToExecSummary("OpWorkingCapital", "LightGreen", "The company's Operational Working Capital is relatively stable. ");
            }
            else if (NoplatResult[1] == "-")
            {
                NegativeComment.Add(NoplatResult[0]);
                color.Add(2);

                AddToExecSummary("OpWorkingCapital", "LightGreen", "The company's Operational Working Capital is decreasing. ");
            }
            else if (NoplatResult[1] == "--")
            {
                NegativeComment.Add(NoplatResult[0]);
                color.Add(3);

                AddToExecSummary("OpWorkingCapital", "Yellow", "The company's Operational Working Capital at a fast pace. What is the reason for this decrease? ");
            }

            if (NoplatResult[2] == "High volatility")
            {
                PositiveComment.Add("High operational working capital volatility");
                color.Add(4);

                AddToExecSummary("OpWorkingCapital", "Red", "The company's operational working capital registered a high volatility. It is important to understand why this is happening.  ");
            }


            //Goodwill
            if (Goodwill.Average() > TotalAssets.Average() * 0.4)
            {
                NegativeComment.Add("Goodwill represents a significant share of total assets. ");
                color.Add(5);

                AddToExecSummary("Goodwill", "Red", "Goodwill represents a significant share of total assets. ");
            }
            else if (Goodwill.Average() > TotalAssets.Average() * 0.2)
            {
                NegativeComment.Add("Goodwill represents a large share of total assets. ");
                color.Add(3);

                AddToExecSummary("Goodwill", "Yellow", "Goodwill represents a large share of total assets. ");
            }
            else if (Goodwill.Average() > TotalAssets.Average() * 0.1)
            {
                PositiveComment.Add("Goodwill represents a reasonable share of total assets. ");
                color.Add(2);

                AddToExecSummary("Goodwill", "LightGreen", "Goodwill represents a reasonable share of total assets. ");
            }
            else if (Goodwill.Average() > 0)
            {
                PositiveComment.Add("Goodwill represents a small share of total assets. ");
                color.Add(2);
                AddToExecSummary("Goodwill", "DarkGreen", "Goodwill represents a small share of total assets. ");
            }
            else
            {
                if (GoodwillRef.Count()==1)
                {
                    NegativeComment.Add("In " + years[GoodwillRef[0]] + " goodwill was" + Goodwill[GoodwillRef[0]] / TotalAssets[GoodwillRef[0]] + "of total assets. ");
                    color.Add(3);

                    AddToExecSummary("Goodwill", "Yellow", "In " + years[GoodwillRef[0]] + " goodwill was" + Goodwill[GoodwillRef[0]] / TotalAssets[GoodwillRef[0]] + "of total assets. ");
                }
                else if (GoodwillRef.Count() > 1)
                {
                    List<string> commentAux = new List<string>();
                    for (int i = 0; i < GoodwillRef.Count(); i++)
                    {
                        if (i==0)
                        {
                            string aux = " " + years[GoodwillRef[i]] ;
                            commentAux.Add(aux);
                        }
                        else if (i == GoodwillRef.Count()-1)
                        {
                            string aux = "and "+ years[GoodwillRef[i]];
                            commentAux.Add(aux);
                        }
                        else
                        {
                            string aux = ", " + years[GoodwillRef[i]];
                            commentAux.Add(aux);
                        }

                    }
                    NegativeComment.Add("In " + commentAux + " goodwill was above 30% of total assets. ");
                    color.Add(3);

                    AddToExecSummary("Goodwill", "Yellow", "In " + commentAux + " goodwill was above 30% of total assets. ");
                }
            }

            Comment(worksheet, 13, column, color, PositiveComment, NegativeComment, OtherComment);
            worksheet.Column(12).Width = 10;
        }
        public void CommentBuilderCFS(ExcelWorksheet worksheet, int row, int column, List<int> years, List<double> FreeCashFlow,
            List<double> CashflowFromOperations, List<double> CashFlowFromFinancing, List<double> CashflowFromInvesting)
        {
            List<string> PositiveComment = new List<string>();
            List<string> NegativeComment = new List<string>();
            List<string> OtherComment = new List<string>();
            List<int> color = new List<int>();

            CalculationsAux values = new CalculationsAux();
            List<int> FreeCashFlowNegative = values.FindNegativeValues(years, FreeCashFlow);
            



            OutliersDetection outliers = new OutliersDetection();
            Dictionary<string, List<int>> OperationsDict = outliers.Outliers(CashflowFromOperations);

            List<double> CashFromInvestingAndFinancing = new List<double>();
            List<double> xvalues = new List<double>();
            for (int i = 0; i < FreeCashFlow.Count(); i++)
            {
                xvalues.Add(i);


            }
            LinearRegression LR = new LinearRegression();
            List<double> LRCashFromOperations = LR.LinearRegressionCalculation(xvalues, CashflowFromOperations);
            List<double> LRfreeCashFlow = LR.LinearRegressionCalculation(xvalues, FreeCashFlow);

            CashflowFromOperations.Reverse();
            List<int> OperationsNegative = values.FindNegativeValues(years, CashflowFromOperations);

            if (OperationsNegative.Count()>0)
            {
                if (OperationsNegative.Count() > 1)
                {
                    
                    NegativeComment.Add("Negative Cash flow from operations in multiple years. ");
                    color.Add(5);

                    AddToExecSummary("CashFromOperations", "Red", "Negative Cash flow from operations in multiple years. ");
                }
                else
                {
                    
                    NegativeComment.Add("Negative Cash flow from operations in " + OperationsNegative[0] + ". ");
                    color.Add(4);

                    AddToExecSummary("CashFromOperations", "Orange", "Negative Cash flow from operations in " + OperationsNegative[0] + ". ");
                }
                
            }

            if (LRCashFromOperations[2]>0.7) //Trend
            {
                if (LRCashFromOperations[0]/Math.Abs(CashflowFromOperations.Average())>0.3)
                {
                    PositiveComment.Add("Cash from operations increased at a fast pace. ");
                    color.Add(1);

                    AddToExecSummary("CashFromOperations", "DarkGreen", "Cash from operations increased at a fast pace. ");
                }
                else if (LRCashFromOperations[0] / Math.Abs(CashflowFromOperations.Average()) > 0.05)
                {
                    PositiveComment.Add("Cash from operations increased. ");
                    color.Add(2);

                    AddToExecSummary("CashFromOperations", "LightGreen", "Cash from operations increased. ");
                }
                else if (LRCashFromOperations[0] / Math.Abs(CashflowFromOperations.Average()) > 0)
                {
                    PositiveComment.Add("Cash from operations stable. ");

                }
                else if (LRCashFromOperations[0] / Math.Abs(CashflowFromOperations.Average()) > -0.1)
                {
                    NegativeComment.Add("Cash from operations declined. ");
                    color.Add(3);

                    AddToExecSummary("CashFromOperations", "Yellow", "Cash from operations declined. ");
                }
                else if (LRCashFromOperations[0] / Math.Abs(CashflowFromOperations.Average()) > -0.3)
                {
                    NegativeComment.Add("Cash from operations declined at a fast pace. ");
                    color.Add(4);

                    AddToExecSummary("CashFromOperations", "Orange", "Cash from operations declined at a fast pace. ");
                }
                else if (LRCashFromOperations[0] / Math.Abs(CashflowFromOperations.Average()) <= -0.3)
                {
                    NegativeComment.Add("Cash from operations declined at an unsettling pace. ");
                    color.Add(5);

                    AddToExecSummary("CashFromOperations", "Red", "Cash from operations declined at an unsettling pace. ");
                }
            }
            else if (LRCashFromOperations[2] > 0.4)
            {
                if (LRCashFromOperations[0] / Math.Abs(CashflowFromOperations.Average()) > 0.5)
                {
                    PositiveComment.Add("Cash from operations increased at a fast pace. ");
                    color.Add(1);

                    AddToExecSummary("CashFromOperations", "DarkGreen", "Cash from operations increased at a fast pace. ");
                }
                else if (LRCashFromOperations[0] / Math.Abs(CashflowFromOperations.Average()) > 0.15)
                {
                    PositiveComment.Add("Cash from operations increased. ");
                    color.Add(2);

                    AddToExecSummary("CashFromOperations", "LightGreen", "Cash from operations increased. ");
                }
                else if (LRCashFromOperations[0] / Math.Abs(CashflowFromOperations.Average()) > 0)
                {

                    PositiveComment.Add("Cash from operations stable. ");

                }
                else if (LRCashFromOperations[0] / Math.Abs(CashflowFromOperations.Average()) > -0.15)
                {
                    NegativeComment.Add("Cash from operations declined. ");
                    color.Add(3);
                    AddToExecSummary("CashFromOperations", "Yellow", "Cash from operations declined. ");
                }
                else if (LRCashFromOperations[0] / Math.Abs(CashflowFromOperations.Average()) > -0.4)
                {
                    NegativeComment.Add("Cash from operations declined at a fast pace. ");
                    color.Add(4);

                    AddToExecSummary("CashFromOperations", "Orange", "Cash from operations declined at a fast pace. ");
                }
                else if (LRCashFromOperations[0] / Math.Abs(CashflowFromOperations.Average()) <= -0.4)
                {
                    NegativeComment.Add("Cash from operations declined at an unsettling pace. ");
                    color.Add(5);

                    AddToExecSummary("CashFromOperations", "Red", "Cash from operations declined at an unsettling pace. ");
                }
            }
            else
            {
                NegativeComment.Add("Cash from operations volatile. ");
                color.Add(3);

                AddToExecSummary("CashFromOperations", "Yellow", "Cash from operations volatile. ");

                if (LRCashFromOperations[0] / Math.Abs(CashflowFromOperations.Average()) > 0.6)
                {
                    PositiveComment.Add("Cash from operations increased at a fast pace. ");
                    color.Add(2);

                    AddToExecSummary("CashFromOperations", "LightGreen", "Cash from operations increased at a fast pace. ");
                }
                else if (LRCashFromOperations[0] / Math.Abs(CashflowFromOperations.Average()) < -0.5)
                {
                    NegativeComment.Add("Cash from operations declined at an unsettling pace. ");
                    color.Add(5);

                    AddToExecSummary("CashFromOperations", "Red", "Cash from operations declined at an unsettling pace. ");
                }
            }

            //if (OperationsDict["PeakNegative"].Count() ==1)
            //{
            //    int aux = OperationsDict["PeakNegative"].FirstOrDefault();
            //    NegativeComment.Add("Significant decrease in Cash from operations in " + years[aux] + ". ");
            //    color.Add(3);
            //}
            //if (OperationsDict["OutlierNegative"].Count() == 1)
            //{
            //    int aux = OperationsDict["OutlierNegative"].FirstOrDefault();
            //    NegativeComment.Add("Significant decrease in Cash from operations in " + years[aux] + ". ");
            //    color.Add(3);
            //}
            //if (OperationsDict["PeakPositive"].Count() == 1)
            //{
            //    int aux = OperationsDict["PeakPositive"].FirstOrDefault();
            //    PositiveComment.Add("Significant increase in Cash from operations in " + years[aux] + ". ");
            //    color.Add(2);
            //}
            //if (OperationsDict["OutlierPositive"].Count() == 1)
            //{
            //    int aux = OperationsDict["OutlierPositive"].FirstOrDefault();
            //    NegativeComment.Add("Significant increase in Cash from operations in " + years[aux] + ". ");
            //    color.Add(2);
            //}

            if (CashflowFromOperations.Average()>CashFlowFromFinancing.Average() && CashflowFromOperations.Average() > CashflowFromInvesting.Average())
            {
                OtherComment.Add("The company funds most of its activity using cash from operations. ");
            }
            else if (CashFlowFromFinancing.Average()>CashflowFromOperations.Average() && CashFlowFromFinancing.Average()>CashflowFromInvesting.Average())
            {
                OtherComment.Add("The company funds most of its activity using cash from financing. ");
            }
            else if (CashflowFromInvesting.Average()>CashflowFromOperations.Average() && CashflowFromInvesting.Average()>CashFlowFromFinancing.Average())
            {
                NegativeComment.Add("The company funds most of its activity using cash from investing activities (What is being sold? And why?). ");
                color.Add(4);
            }

            if (CashflowFromOperations.Average() + CashflowFromInvesting.Average() + CashFlowFromFinancing.Average()<0)
            {
                NegativeComment.Add("On average, the company is not being able to generate positive cash flow. ");
                color.Add(3);
            }

            Comment(worksheet, 26, column, color, PositiveComment, NegativeComment, OtherComment);
            worksheet.Column(12).Width = 10;
        }

        public List<string> simpleAnalysis(List<double> LinearRegression, double Cagr, string Caption)
        {
            List<string> output = new List<string>();

            if (LinearRegression[2] > 0.5)
            {
                if (Cagr > 0.3)
                {
                    output.Add(Caption + " is rising at a fast pace. ");
                    output.Add("++");
                }
                else if (Cagr > 0.05)
                {
                    output.Add(Caption + " is rising. ");
                    output.Add("+");
                }
                else if (Cagr > -0.05)
                {
                    output.Add(Caption + " is stable. ");
                    output.Add("+-");
                }
                else if (Cagr > -0.2)
                {
                    output.Add(Caption + " is decreasing. ");
                    output.Add("-");
                }
                else
                {
                    output.Add(Caption + " is decreasing at a fast pace. ");
                    output.Add("--");
                }
                output.Add("Low Volatility");

            }
            else
            {
                if (Cagr > 0.4)
                {
                    output.Add(Caption + " is rising at a fast pace. ");
                    output.Add("++");
                }
                else if (Cagr > 0.1)
                {
                    output.Add(Caption + " is rising. ");
                    output.Add("+");
                }
                else if (Cagr > -0.1)
                {
                    output.Add(Caption + " is showing no clear trend. ");
                    output.Add("+-");
                }
                else if (Cagr > -0.3)
                {
                    output.Add(Caption + " is decreasing. ");
                    output.Add("-");
                }
                else
                {
                    output.Add(Caption + " is decreasing at a fast pace. ");
                    output.Add("--");
                }
                output.Add("High Volatility");
            }
            return output;

        }
    }


}
