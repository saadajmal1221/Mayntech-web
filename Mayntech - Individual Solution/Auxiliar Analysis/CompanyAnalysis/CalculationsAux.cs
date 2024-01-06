using Mayntech___Individual_Solution.Auxiliar.Analysis;
using Mayntech___Individual_Solution.Auxiliar.Executive_Summary;
using Mayntech___Individual_Solution.Auxiliar_Analysis.Valuation;

namespace Mayntech___Individual_Solution.Auxiliar.CompanyAnalysis
{
    public class CalculationsAux
    {
        public List<int> FindNegativeValues(List<int> years, List<double> values)
        {
            List<int> negativeYears = new List<int>();

            for (int i = 0; i < values.Count(); i++)
            {
                if (values[i]<0)
                {
                    negativeYears.Add(years[i]);
                }
            }
            return negativeYears;
        }

        public List<int> PercentageRef(List<double> reference, List<double> values, double refValue)
        {
            List<int> result = new List<int>();

            for (int i = 0; i < reference.Count(); i++)
            {
                if (values[i] / reference[i] > refValue)
                {
                    result.Add(i);
                }
            }

            return result;
        }

        public string GrossMarginEvol(List<double> GrossProfit, List<double> GrossMargin, List<double> xvalues)
        {
            ValuationSupport support = new ValuationSupport();
            string result = null;

            double cagr = support.Cagr(GrossMargin);

            LinearRegression linearRegression = new LinearRegression();
            List<double> LRGrossMargin = linearRegression.LinearRegressionCalculation(xvalues, GrossMargin);
            List<double> LRGrossProfit = linearRegression.LinearRegressionCalculation(xvalues, GrossProfit);


            if (LRGrossMargin[2] > 0.7)
            {
                if (LRGrossMargin[0]/GrossMargin.Average()>0.15)
                {
                    result = "Fast Increasing Rate";
                }
                else if (LRGrossMargin[0] / GrossMargin.Average() > 0.02)
                {
                    result = "Increasing Rate";
                }
                else if (LRGrossMargin[0] / GrossMargin.Average() > -0.01)
                {
                    result = "Same Rate";
                }
                else if(cagr<0)
                {

                    if (LRGrossProfit[2]>0.5)
                    {
                        if (LRGrossProfit[0]/GrossProfit.Average()<0)
                        {
                            result = "Decreasing Absolute";
                        }
                        else
                        {
                            result = "Decreasing Rate";
                        }
                    }
                    else
                    {
                        if (LRGrossProfit[0] / GrossProfit.Average() < -0.1)
                        {
                            result = "Decreasing Absolute";
                        }
                        else
                        {
                            result = "Decreasing Rate";
                        }
                    }
                }
            }
            else if (LRGrossMargin[2] > 0.4)
            {
                if (LRGrossMargin[0] / GrossMargin.Average() > 0.25)
                {
                    result = "Fast Increasing Rate";
                }
                else if (LRGrossMargin[0] / GrossMargin.Average() > 0.1)
                {
                    result = "Increasing Rate";
                }
                else if (LRGrossMargin[0] / GrossMargin.Average() > -0.04)
                {
                    result = "Same Rate";
                }
                else if (cagr < 0)
                {
                    if (LRGrossProfit[2] > 0.5)
                    {
                        if (LRGrossProfit[0] < 0)
                        {
                            result = "Decreasing Absolute";
                        }
                        else
                        {
                            result = "Decreasing Rate";
                        }
                    }
                    else
                    {
                        if (LRGrossProfit[0] < -0.1)
                        {
                            result = "Decreasing Absolute";
                        }
                        else
                        {
                            result = "Decreasing Rate";
                        }
                    }
                }
            }
            else
            {
                if (LRGrossMargin[0] > 0.3 && cagr>0)
                {
                    result = "Fast Increasing Rate";
                }

                else if (LRGrossMargin[0] / GrossMargin.Average() > -0.07)
                {
                    result = "Volatile";
                }
                else if (cagr < 0)
                {
                    if (LRGrossProfit[2] > 0.5)
                    {
                        if (LRGrossProfit[0] < 0)
                        {
                            result = "Decreasing Absolute";
                        }
                        else
                        {
                            result = "Decreasing Rate";
                        }
                    }
                    else
                    {
                        if (LRGrossProfit[0] < -0.1)
                        {
                            result = "Decreasing Absolute";
                        }
                        else
                        {
                            result = "Decreasing Rate";
                        }
                    }
                }
            }
            return result;
        }

        public string IsGrowing(List<double> LRvalues)
        {
            string result = null;
            if (LRvalues[2]>0.5)
            {
                if (LRvalues[0]>0)
                {
                    result = "Yes";
                }
                else
                {
                    result = "No";
                }
            }
            else
            {
                if (LRvalues[0] > 0.05)
                {
                    result = "Yes";
                }
                else
                {
                    result = "No";
                }
            }
            return result;
        }
        public string IsDecreasing(List<double> LRvalues)
        {
            string result = null;
            if (LRvalues[2] > 0.5)
            {
                if (LRvalues[0] < 0)
                {
                    result = "Yes";
                }
                else
                {
                    result = "No";
                }
            }
            else
            {
                if (LRvalues[0] < -0.05)
                {
                    result = "Yes";
                }
                else
                {
                    result = "No";
                }
            }
            return result;
        }

        public void AddToExecSummary(string Topic, string color, string description)
        {

            List<string> details = new List<string>();

            details.AddRange(new List<string> { Topic, color, description });

            ExecSummary.ExecSummaryDetails.Add( details);
        }

    }
}
