using static OfficeOpenXml.ExcelErrorValue;

namespace Mayntech___Individual_Solution.Auxiliar.CompanyAnalysis
{
    public class OutliersDetection
    {
        public Dictionary<string, List<int>> Outliers(List<double> values)
        {
            //Find outliers
            IDictionary<string, List<int>> outliers = new Dictionary<string, List<int>>();
            //outliers.Add("Outlier", null);
            //outliers.Add("Peak", null);

            List<int> outliersNegativeCount = new List<int>();
            List<int> outliersPositiveCount = new List<int>();
            List<int> peakPositiveCount = new List<int>();
            List<int> peakNegativeCount = new List<int>();

            for (int i = 0; i < values.Count(); i++)
            {
                if (i == 0)
                {
                    double difFut = (values[i] - values[i + 1]) / values.Average();
                    if (difFut < -1.5)
                    {
                        
                    }
                    else if (difFut > 0.3)
                    {
                        
                    }
                }
                else if (i != values.Count() - 1)
                {
                    double difPast = (values[i] - values[i - 1]) / values.Average();
                    double difFut = (values[i] - values[i + 1]) / values.Average();
                    if (difPast > 0.3 && difFut > 0.3)
                    {
                        outliersPositiveCount.Add(i);
                    }
                    else if (difPast < -0.3 && difFut < -0.3)
                    {
                        outliersNegativeCount.Add(i);
                    }
                    else if (difPast > 1.5)
                    {
                        peakPositiveCount.Add(i);
                    }
                    else if (difPast < -1.5)
                    {
                        peakNegativeCount.Add(i);
                    }
                }
                else
                {
                    double difPast = (values[i] - values[i - 1]) / values.Average();
                    if (difPast < -1.5)
                    {
                        peakNegativeCount.Add(i);
                    }
                    else if (difPast > 1.5)
                    {
                        peakPositiveCount.Add(i);
                    }
                }
                
            }
            outliers.Add("OutlierPositive", outliersPositiveCount);
            outliers.Add("OutlierNegative", outliersNegativeCount);
            outliers.Add("PeakPositive", peakPositiveCount);
            outliers.Add("PeakNegative", peakNegativeCount);

            return (Dictionary<string, List<int>>)outliers;
        }
    }
}
