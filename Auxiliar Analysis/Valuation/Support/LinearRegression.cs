using System.Diagnostics;

namespace Mayntech___Individual_Solution.Auxiliar.Analysis
{
    public class LinearRegression
    {
        public List<double> LinearRegressionCalculation(List<double> xvalues, List<double> yvalues)
        {
            

            if (xvalues.Count()>yvalues.Count())
            {
                xvalues = RemovesUntilEqual(xvalues, yvalues)[0];
            }
            else if (yvalues.Count() > xvalues.Count())
            {
                yvalues = RemovesUntilEqual(xvalues, yvalues)[1];
            }
            
            double varianceOfX = 0;
            double covarianceOfXandY = 0;
            double SumOfRegression = 0;
            double SumOfTotal = 0;
            double Rsquared = 0;


            for (int i = 0; i < xvalues.Count(); i++)
            {
                double x = xvalues[i];
                double y = yvalues[i];
                varianceOfX += (xvalues.Average() - x) * (xvalues.Average() - x);
                covarianceOfXandY += (yvalues.Average() - y) * (xvalues.Average() - x);
            }
            double slope = covarianceOfXandY / varianceOfX;
            double intercept = yvalues.Average() - slope * xvalues.Average();

            List<double> outputLinearRegression = new List<double>();
            outputLinearRegression.Add(slope);
            outputLinearRegression.Add(intercept);

            //Calculate the R squared
            for (int i = 0; i < xvalues.Count(); i++)
            {
                SumOfRegression += (yvalues[i] - (intercept + (slope * xvalues[i]))) * (yvalues[i] - (intercept + (slope * xvalues[i])));
                SumOfTotal += (yvalues[i] - yvalues.Average()) * (yvalues[i] - yvalues.Average());
            }
            Rsquared = 1 - (SumOfRegression / SumOfTotal);
            outputLinearRegression.Add(Rsquared);

            return outputLinearRegression;

        }

        public List<List<double>> RemovesUntilEqual(List<double> xvalues, List<double> yvalues)
        {
            List<List<double>> output = new List<List<double>>();

            if (xvalues.Count() > yvalues.Count())
            {
                xvalues.RemoveAt(xvalues.Count() - 1);
                return RemovesUntilEqual(xvalues, yvalues);
            }
            else if (xvalues.Count() < yvalues.Count())
            {
                yvalues.RemoveAt(xvalues.Count() - 1);
                return RemovesUntilEqual(xvalues, yvalues);
            }
            else
            {
                output.Add(xvalues);
                output.Add(yvalues);
                return output;
            }
        }
    }
}
