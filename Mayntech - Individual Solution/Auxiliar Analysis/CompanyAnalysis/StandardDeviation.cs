using Mayntech___Individual_Solution.Auxiliar.Analysis;
using static OfficeOpenXml.ExcelErrorValue;

namespace Mayntech___Individual_Solution.Auxiliar.CompanyAnalysis
{
    public class StandardDeviation
    {
        public double Sd(List<double> values)
        {
            double average = values.Average();
            double sumOfSquaresOfDifferences = values.Select(val => (val - average) * (val - average)).Sum();
            double sd = Math.Sqrt(sumOfSquaresOfDifferences / values.Count());

            return sd;
        }
    }
}
