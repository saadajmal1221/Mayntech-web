using Mayntech___Individual_Solution.Auxiliar.Auxiliar.FMP;
using Mayntech___Individual_Solution.Auxiliar.Executive_Summary;

namespace Mayntech___Individual_Solution.Auxiliar_Analysis.CompetitorsRatios
{
    public class CommentsRatios
    {
        public string CommentAvrgRatio(string companyTick,
            IDictionary<string, double> AllCompaniesAverage)
        {
            string output = null;
            double avrg = AllCompaniesAverage[companyTick];
            double avrgTotal = AllCompaniesAverage.Values.Average();

            var sortedAverageDict = from entry in AllCompaniesAverage orderby entry.Value descending select entry;

            KeyValuePair<string, double> minimumValue = sortedAverageDict.Last();
            int number = sortedAverageDict.Count();

            try
            {
                KeyValuePair<string, double> maximumValue = sortedAverageDict.First();
                KeyValuePair<string, double> SecondMax = sortedAverageDict.ElementAt(1);
                KeyValuePair<string, double> SecondMin = sortedAverageDict.ElementAt(sortedAverageDict.Count() - 2);
                double secondMax = SecondMax.Value;
                double secondMin = SecondMin.Value;
                double min = minimumValue.Value;
                double max = maximumValue.Value;

                if (avrg > avrgTotal)
                {
                    output += "+";
                    if (avrg > 0.95 * max)
                    {
                        if (avrg > 5 * secondMax)
                        {
                            output += "!+";
                        }
                        else if (avrg > 2 * secondMax)
                        {
                            output += "++";
                        }
                        else
                        {
                            output += "+";
                        }

                    }
                }
                else if (avrg <= avrgTotal)
                {
                    output += "-";
                    if (avrg < 1.05 * min)
                    {
                        if (avrg < 0.05 * secondMin)
                        {
                            output += "!-";
                        }
                        else if (avrg < 0.5 * secondMin)
                        {
                            output += "--";
                        }
                        else
                        {
                            output += "-";
                        }

                    }
                }
            }
            catch 
            {

                return null;
            }

            return output;
        }


        public string CommentSdRatio(string companyTick,
            IDictionary<string, double> AllCompaniesSd)
        {
            string output = null;
            double avrg = AllCompaniesSd[companyTick];
            double avrgTotal = AllCompaniesSd.Values.Average();
            

            var sortedSdDict = from entry in AllCompaniesSd orderby entry.Value descending select entry;

            KeyValuePair<string, double> minimumValue = sortedSdDict.Last();
            KeyValuePair<string, double> maximumValue = sortedSdDict.First();
            
            double min = minimumValue.Value;
            double max = maximumValue.Value;
            

            if (avrg > avrgTotal)
            {
                output += "+";
                if (avrg > 0.90 * max)
                {
                    output += "+";

                }
            }
            else if (avrg <= avrgTotal)
            {
                output += "-";
                if (avrg < 1.1 * min)
                {
                    output += "--";
                }
            }
            return output;
        }

        public double StandardDeviation(List<double> values)
        {
            double MainCompanyAverage = values.Average();
            double MainsumOfSquaresOfDifferences = values.Select(val => (val - MainCompanyAverage) * (val - MainCompanyAverage)).Sum();
            double Mainsd = Math.Sqrt(MainsumOfSquaresOfDifferences / values.Count());

            return Mainsd;
        }

        public double PointCalculator(IDictionary<string, double> AllCompaniesAverage, IDictionary<string, double> AllCompaniesSd, string companyTick, int Nature, List<string> comments, string ratio)
        {
            string outputAvrg = CommentAvrgRatio(companyTick, AllCompaniesAverage);
            string outputSd = CommentSdRatio(companyTick, AllCompaniesSd);
            

            double points = 0;

            //Standard deviation
            if (outputSd == "--")
            {
                points += 0.25;
            }
            else if (outputSd == "-")
            {
                points += 0.15;
            }
            else if (outputSd == "+")
            {
                points+= 0.05;
            }
            else
            {
                points+=0;
            }

            //Average
            if (outputAvrg == "--")
            {
                if (Nature == 1) //Nature: [1 (alto), 2 (Médio), 3 (baixo)]
                {
                    points += 0;
                    //ExecSummary.ExecSummaryDetails.Add( new List<string> {  ratio, "Yellow", comments[1] });

                }
                else if (Nature == 3)
                {
                    points += 0.75;
                }
                else
                {
                    points += 0.4;
                }
                
            }
            else if (outputAvrg == "-!-")
            {
                //ExecSummary.ExecSummaryDetails.Add(new List<string> {  ratio, "Yellow", comments[2] });
                points += 0;
                if (Nature == 3)
                {
                    points += 0.2;
                }
            }
            else if (outputAvrg == "-")
            {
                if (Nature == 1)
                {
                    points += 0.2;
                }
                else if (Nature == 3)
                {
                    points += 0.7;
                }
                else
                {
                    points += 0.75;
                }
                
            }
            else if (outputAvrg == "+")
            {
                if (Nature == 1)
                {
                    points += 0.7;
                }
                else if (Nature == 3)
                {
                    points += 0.2;
                }
                else
                {
                    points += 0.75;
                }
                
            }
            else if (outputAvrg == "++")
            {
                if (Nature == 1)
                {
                    points += 0.75;


                }
                else if (Nature == 3)
                {
                    points += 0;
                    //ExecSummary.ExecSummaryDetails.Add(new List<string> {  ratio, "Yellow", comments[1] });
                }
                else
                {
                    points += 0.4;
                }
                
            }
            else if (outputAvrg == "+++")
            {
                if (Nature == 1)
                {
                    points += 0.5;
                }
                else if (Nature == 3)
                {
                    points += 0;
                }
                else
                {
                    points += 0.2;
                }
            }
            else if (outputAvrg == "---")
            {
                if (Nature == 1)
                {
                    points += 0;
                }
                else if (Nature == 3)
                {
                    points += 0.5;
                }
                else
                {
                    points += 0.2;
                }
            }
            else
            {
                //ExecSummary.ExecSummaryDetails.Add(new List<string> {  ratio, "Yellow", comments[2] });
                if (Nature == 1)
                {
                    points += 0.2;
                }
                points += 0;
            }

            return points;
        }

    }
}
