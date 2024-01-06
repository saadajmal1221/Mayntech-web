using Mayntech___Individual_Solution.Auxiliar.Analysis;
using Mayntech___Individual_Solution.Auxiliar.Auxiliar.Other;
using Mayntech___Individual_Solution.Auxiliar.Executive_Summary;
using Mayntech___Individual_Solution.Pages;
using OfficeOpenXml;
using OfficeOpenXml.Filter;
using OfficeOpenXml.Sparkline;
using OfficeOpenXml.Style;
using System;
using System.Data.Common;
using System.Drawing;
using System.Runtime.Intrinsics.X86;
using static OfficeOpenXml.ExcelErrorValue;

namespace Mayntech___Individual_Solution.Auxiliar.CompanyAnalysis.Competitors
{
    public class CompetitorsAux
    {
        public List<List<string>> SupportCompetitors(ExcelWorksheet worksheet, int row, int col, List<double> peers, List<double> reference, string evolution,
            string size, int rowNumb, string Caption, string companyName, string nature, List<string> commentList, 
            List<int> auxiliar, string coefficient, List<string> commentListSize, List<string> commentListCoeficient, int numberOfColumns)
        {
            ExcelNextCol columnName = new ExcelNextCol();

            for (int i = 0; i < numberOfColumns +4; i++)
            {
                if (Caption == "Revenue")
                {
                    if (i!=numberOfColumns+1)
                    {
                        worksheet.Cells[row + 3, col + i].Style.Fill.PatternType = ExcelFillStyle.Solid;
                        worksheet.Cells[row + 3, col + i].Style.Fill.BackgroundColor.SetColor(Cores.corSecundária);
                        worksheet.Cells[row + 4, col + i].Style.Fill.PatternType = ExcelFillStyle.Solid;
                        worksheet.Cells[row + 4, col + i].Style.Fill.BackgroundColor.SetColor(Cores.corSecundária);
                    }

                }
                else
                {
                    if (i != numberOfColumns+1)
                    {
                        worksheet.Cells[row + 3, col + i].Style.Fill.PatternType = ExcelFillStyle.Solid;
                        worksheet.Cells[row + 3, col + i].Style.Fill.BackgroundColor.SetColor(Cores.CorNumber3);
                        worksheet.Cells[row + 4, col + i].Style.Fill.PatternType = ExcelFillStyle.Solid;
                        worksheet.Cells[row + 4, col + i].Style.Fill.BackgroundColor.SetColor(Cores.CorNumber3);
                    }

                }
                if (i == 0)
                {
                    worksheet.Cells[row + 1, numberOfColumns+8].Value = "Comments";
                    worksheet.Cells[row + 1, numberOfColumns + 8].Style.Font.Bold = true;
                    worksheet.Cells[row + 1, numberOfColumns + 8].Style.Border.Bottom.Style = ExcelBorderStyle.Thin;
                    worksheet.Columns[numberOfColumns + 8].Width = 55;

                    int startRow = row + 3;
                    int endRow = row + 5;
                    string CommentCol = columnName.GetExcelColumnName(numberOfColumns + 8);
                    worksheet.Cells[CommentCol + startRow + ":"+ CommentCol + endRow].Merge = true;
                    worksheet.Cells[CommentCol + startRow].Style.HorizontalAlignment = ExcelHorizontalAlignment.Left;
                    worksheet.Cells[CommentCol + startRow].Style.VerticalAlignment = ExcelVerticalAlignment.Top;
                    worksheet.Cells[CommentCol + startRow].Style.WrapText = true;

                    worksheet.Cells[row + 1, numberOfColumns + 7].Value = "Indicator";
                    worksheet.Cells[row + 1, numberOfColumns + 7].Style.Font.Bold = true;
                    worksheet.Cells[row + 1, numberOfColumns + 7].Style.Border.Bottom.Style = ExcelBorderStyle.Thin;
                    worksheet.Columns[numberOfColumns + 7].Width = 10;


                    worksheet.Cells[row+1, col+i].Value = Caption;
                    worksheet.Cells[row+1, col+i].Style.Font.Bold = true;
                    worksheet.Cells[row+1, col + i].Style.Font.Color.SetColor(Cores.CorTexto);

                    worksheet.Row(row + 2).Height = 4;


                    worksheet.Cells[row+3, col + i].Value = companyName;
                    worksheet.Cells[row + 3, col + i].Style.Indent = 2;

                    worksheet.Cells[row + 4, col + i].Value = "Competitors";
                    worksheet.Cells[row + 4, col + i].Style.Indent = 2;
                    worksheet.Columns[col + i].Width = 35;
                }
                else if (i >0 && i<numberOfColumns+1)
                {
                    int aux = SolutionModel.NumberYears - (numberOfColumns-1) + 1 + i;
                    string column = columnName.GetExcelColumnName(aux);

                    try
                    {
                        if (Caption == "Revenue" || Caption == "Gross Margin" || Caption == "Operating income Margin" || Caption == "EBITDA Margin" || Caption == "Net Income Margin")
                        {


                            if (Caption == "Revenue")
                            {



                                worksheet.Cells[row + 4, col + i].Value = peers[i - 1] / 1000;
                                worksheet.Cells[row + 4, col + i].Style.Numberformat.Format = "#,##0;(#,##0);-";

                                worksheet.Cells[row + 3, col + i].Formula = "='P&L'!" + column + rowNumb;
                                worksheet.Cells[row + 3, col + i].Style.Numberformat.Format = "#,##0;(#,##0);-";



                                worksheet.Columns[col + i].Width = 15;
                            }
                            else
                            {
                                worksheet.Cells[row + 3, col + i].Formula = "='P&L'!" + column + rowNumb;
                                worksheet.Cells[row + 3, col + i].Style.Numberformat.Format = "0.00%";


                                worksheet.Cells[row + 4, col + i].Value = peers[i - 1];
                                worksheet.Cells[row + 4, col + i].Style.Numberformat.Format = "0.00%";

                                worksheet.Columns[col + i].Width = 15;
                            }
                        }
                        else
                        {
                            worksheet.Cells[row + 3, col + i].Formula = "='BS'!" + column + rowNumb + "/'BS'!" + column + "23";
                            worksheet.Cells[row + 3, col + i].Style.Numberformat.Format = "0.00%";


                            worksheet.Cells[row + 4, col + i].Value = peers[i - 1];
                            worksheet.Cells[row + 4, col + i].Style.Numberformat.Format = "0.00%";

                            worksheet.Columns[col + i].Width = 15;
                        }
                    }
                    catch (Exception)
                    {

                        
                    }

                    


                }
                else if (i == numberOfColumns+1)
                {
                    worksheet.Columns[col + i].Width = 2;
                }
                else if (i==numberOfColumns+2)
                {
                    int auxcol = numberOfColumns + 2;
                    int sparkAux = numberOfColumns + 6;
                    int cagrAux = numberOfColumns - 1;
                    string Lastcolumn = columnName.GetExcelColumnName(auxcol);
                    string SparkColumn = columnName.GetExcelColumnName(sparkAux);

                    for (int a = 0; a < 2; a++)
                    {
                        int rowFormula = row + 3 + a;
                        string aux = '"' + "n.a." + '"';
                        if (peers.Count()<5)
                        {
                            int yearForFormulas = peers.Count() - 1;
                        }
                        worksheet.Cells[rowFormula, col + i].Formula = "=IFERROR(IF(AND(C" + rowFormula + "<0," + Lastcolumn + rowFormula + "<0),-(("+ Lastcolumn + rowFormula + "/C" + rowFormula + ")^(1/" + cagrAux + ")-1),("+ Lastcolumn + rowFormula + "/C" + rowFormula + ")^(1/" + cagrAux + ")-1)," + aux + ")";
                        worksheet.Cells[rowFormula, col + i].Style.Numberformat.Format = "0.00%";
                        var sparklineLine = worksheet.SparklineGroups.Add(eSparklineType.Line, worksheet.Cells[SparkColumn + rowFormula], worksheet.Cells["C" + rowFormula + ":"+ Lastcolumn + rowFormula]);
                        worksheet.Column(col + i).Style.HorizontalAlignment = ExcelHorizontalAlignment.Right;
                    }
                    worksheet.Columns[col + i].Width = 10;
                }
                else if (i==numberOfColumns+3)
                {
                    int auxcol = numberOfColumns + 2;
                    int sparkAux = numberOfColumns + 6;
                    int cagrAux = numberOfColumns - 1;
                    string Lastcolumn = columnName.GetExcelColumnName(auxcol);
                    string SparkColumn = columnName.GetExcelColumnName(sparkAux);
                    for (int a = 0; a < 2; a++)
                    {
                        int rowFormula = row + 3 + a;
                        string aux = '"' + "n.a." + '"';
                        worksheet.Cells[rowFormula, col + i].Formula = "=IFERROR(IF(AND(C" + rowFormula + "<0," + Lastcolumn + rowFormula + "<0),-((" + Lastcolumn + rowFormula + "/C" + rowFormula + ")^(1/" + cagrAux + ")-1),(" + Lastcolumn + rowFormula + "/C" + rowFormula + ")^(1/" + cagrAux + ")-1)," + aux + ")";
                        worksheet.Cells[rowFormula, col + i].Style.Numberformat.Format = "0.00%";
                        worksheet.Column(col + i).Style.HorizontalAlignment = ExcelHorizontalAlignment.Right;
                    }
                    worksheet.Columns[col + i].Width = 24;
                }

            }
            List<List<string>> output = new List<List<string>>();
            
            output = comments(worksheet, evolution, size, row + 3, nature, commentList, Caption, auxiliar, coefficient, commentListSize, commentListCoeficient, peers, reference, numberOfColumns+7);
            return output;
        }

        public string Evolution(List<double> reference, List<double> peers)
        {
            string result = null;
            List<double> diference = new List<double>();
            List<double> xvalues = new List<double>();
            for (int i = 0; i < reference.Count(); i++)
            {
                diference.Add(reference[i]- peers[i]);
                xvalues.Add(i);
            }

            LinearRegression linearRegression = new LinearRegression();
            List<double> LrOutput = linearRegression.LinearRegressionCalculation(xvalues, diference);
            List<double> LrReference = linearRegression.LinearRegressionCalculation(xvalues, reference);
            List<double> LrPeers = linearRegression.LinearRegressionCalculation(xvalues, peers);

            double teste = LrOutput[0] / reference.Average();
            double cagrRef = CAGR(reference);
            double cagrPeers = CAGR(peers);
            

            if (double.IsNaN(cagrPeers) != true && double.IsNaN(cagrRef)!=true)
            {
                if (LrReference[2] > 0.5 && LrPeers[2] > 0.5)
                {
                    if (cagrRef > 1.5 * cagrPeers)
                    {
                        result = "++";

                    }
                    else if (cagrRef > 1.1 * cagrPeers)
                    {
                        result = "+";
                    }
                    else if (cagrRef > 0.9 * cagrPeers)
                    {
                        result = "+-";
                    }
                    else if (cagrRef > 0.5 * cagrPeers)
                    {
                        result = "-";
                    }
                    else
                    {
                        result = "--";
                    }
                }
                else if (LrReference[2] > 0.5)
                {
                    if (cagrRef > 2 * cagrPeers)
                    {
                        result = "++";

                    }
                    else if (cagrRef > 1.3 * cagrPeers)
                    {
                        result = "+";
                    }
                    else if (cagrRef > 0.7 * cagrPeers)
                    {
                        result = "+-";
                    }
                    else if (cagrRef > 0.5 * cagrPeers)
                    {
                        result = "-";
                    }
                    else
                    {
                        result = "--";
                    }
                }
                else if (LrPeers[2] > 0.7)
                {
                    if (cagrRef > 2 * cagrPeers)
                    {
                        result = "++";

                    }
                    else if (cagrRef > 1.3 * cagrPeers)
                    {
                        result = "+";
                    }
                    else if (cagrRef > 0.7 * cagrPeers)
                    {
                        result = "+-";
                    }
                    else if (cagrRef > 0.5 * cagrPeers)
                    {
                        result = "-";
                    }
                    else
                    {
                        result = "--";
                    }
                }
                else
                {
                    if (cagrRef* cagrPeers<0)
                    {
                        if (cagrRef>0)
                        {
                            result = "+";
                        }
                        else
                        {
                            result = "-";
                        }
                    }
                    else
                    {
                        result = "n.a.";
                    }
                    

                }
            }
            else
            {

                if (LrPeers[2] > 0.2 && LrReference[2] > 0.2)
                {
                    double auxiliar = (reference[reference.Count() - 1] / reference[0]) / 4;
                    double auxiliarPeers = (peers[peers.Count() - 1] / peers[0]) / 4;

                    if (auxiliar > 1.3 * auxiliarPeers)
                    {
                        result = "++";
                    }
                    else if (auxiliar > 1.1 * auxiliarPeers)
                    {
                        result = "+";
                    }
                    else if (auxiliar > 0.9 * auxiliarPeers)
                    {
                        result = "+-";
                    }
                    else if (auxiliar > 0.7 * auxiliarPeers)
                    {
                        result = "-";
                    }
                    else
                    {
                        result = "--";
                    }

                }
                else
                {
                    result = "n.a.";
                }
            }
            

            return result;

        }

        public string EvolutionBS(List<double> reference, List<double> peers)
        {
            string result = null;
            List<double> diference = new List<double>();
            List<double> xvalues = new List<double>();
            for (int i = 0; i < reference.Count(); i++)
            {
                diference.Add(reference[i] - peers[i]);
                xvalues.Add(i);
            }

            LinearRegression linearRegression = new LinearRegression();
            List<double> LrOutput = linearRegression.LinearRegressionCalculation(xvalues, diference);
            List<double> LrReference = linearRegression.LinearRegressionCalculation(xvalues, reference);
            List<double> LrPeers = linearRegression.LinearRegressionCalculation(xvalues, peers);

            double teste = LrOutput[0] / reference.Average();
            double cagrRef = CAGR(reference);
            double cagrPeers = CAGR(peers);


            if (double.IsNaN(cagrPeers) != true && double.IsNaN(cagrRef) != true)
            {
                if (cagrRef*cagrPeers<0)
                {
                    if (cagrRef>0 && cagrRef>0.02 || cagrPeers<-0.02)
                    {
                        result = "+";
                    }
                    else if (cagrRef < 0 && cagrRef < -0.02 || cagrPeers > 0.02)
                    {
                        result = "-";
                    }
                    else
                    {
                        result = "+-";
                    }
                }
                else
                {
                    if (Math.Abs(Math.Max(cagrRef, cagrPeers))>0.2 && Math.Abs(Math.Max(cagrRef, cagrPeers)) < 0.5)
                    {
                        if (cagrRef>3*cagrPeers)
                        {
                            result = "++";
                        }
                        else if (cagrRef > 2 * cagrPeers)
                        {
                            result = "+";
                        }
                        else if (cagrRef < 0.5*cagrPeers)
                        {
                            result = "-";
                        }
                        else if(cagrRef < 0.33 * cagrPeers)
                        {
                            result = "--";
                        }
                    }
                    else if (Math.Abs(Math.Max(cagrRef, cagrPeers)) > 0.5)
                    {
                        if (cagrRef > 1.5 * cagrPeers)
                        {
                            result = "++";
                        }
                        else if (cagrRef > 1.4 * cagrPeers)
                        {
                            result = "+";
                        }
                        else if (cagrRef < 0.72 * cagrPeers)
                        {
                            result = "-";
                        }
                        else if (cagrRef < 0.66 * cagrPeers)
                        {
                            result = "--";
                        }
                    }
                    else
                    {
                        if (cagrRef >  cagrPeers+0.1)
                        {
                            result = "+";
                        }
                        else if (cagrRef + 0.1 < cagrPeers )
                        {
                            result = "-";
                        }
                    }


                }
            }
            else
            {

                if (LrPeers[2] > 0.2 && LrReference[2] > 0.2)
                {
                    double auxiliar = (reference[reference.Count() - 1] / reference[0]) / 4;
                    double auxiliarPeers = (peers[peers.Count() - 1] / peers[0]) / 4;

                    if (auxiliar* auxiliarPeers<0)
                    {
                        if (auxiliar>0)
                        {
                            result = "+";
                        }
                        else
                        {
                            result = "-";
                        }
                    }

                }
                else
                {
                    result = "n.a.";
                }
            }


            return result;

        }
        public string Size(List<double> reference, List<double> peers)
        {
            string result = null;


            if (reference.Average()> peers.Average()*1.5)
            {
                result = "++";

            }
            else if (reference.Average() > peers.Average() * 1.05)
            {
                result = "+";
            }
            else if (reference.Average() > peers.Average() * 0.95)
            {
                result = "+-";
            }
            else if (reference.Average() > peers.Average() * 0.5)
            {
                result = "-";
            }
            else 
            {
                result = "--";
            }

            return result;
        }
        public string coefficientOfVariation(List<double> reference, List<double> peers)
        {
            string output = null;

            double average = reference.Average();
            double sumOfSquaresOfDifferences = reference.Select(val => (val - average) * (val - average)).Sum();
            double sdReference = Math.Sqrt(sumOfSquaresOfDifferences / (reference.Count()-1));
            double CoefficientReference = sdReference / reference.Average();

            double averagePeers = peers.Average();
            double sumOfSquaresOfDifferencesPeers = peers.Select(val => (val - averagePeers) * (val - averagePeers)).Sum();
            double sdPeers = Math.Sqrt(sumOfSquaresOfDifferencesPeers / (peers.Count()-1));
            double CoefficientPeers = sdPeers / peers.Average();

            List<double> xvalues = new List<double>();
            for (int i = 0; i < reference.Count(); i++)
            {
                xvalues.Add(i + 1);
            }

            LinearRegression linearRegression = new LinearRegression();
            List<double> RegressionRef = linearRegression.LinearRegressionCalculation(xvalues, reference);

            if (RegressionRef[2]<0.7)
            {
                if (Math.Abs(CoefficientReference) > Math.Abs(3 * CoefficientPeers))
                {
                    output = "++";

                }
                else if (Math.Abs(CoefficientReference) > Math.Abs(1.1 * CoefficientPeers))
                {
                    output = "+";
                }
                else if (Math.Abs(CoefficientReference) > Math.Abs(0.9 * CoefficientPeers))
                {
                    output = "+-";
                }
                else if (Math.Abs(CoefficientReference) > Math.Abs(0.3 * CoefficientPeers))
                {
                    output = "-";
                }
                else
                {
                    output = "--";
                }
                
            }
            else
            {
                output = "n.a.";
            }
            return output;

        }
        public double CAGR(List<double> values)
        {
            double cagr = 0;
            try
            {
                if (values[0] < 0 && values[values.Count() - 1] < 0)
                {
                    cagr = -Math.Pow((values[values.Count() - 1]) / (values[0]), (1 / values.Count())) - 1;
                }
                else
                {
                    double aux6 = 1;
                    
                    cagr = Math.Pow((values[values.Count() - 1]) / (values[0]), (aux6/ (values.Count()-1))) - 1;
                }
                
            }
            catch (Exception)
            {

                return 0;
            }
            return cagr;
        }

        public List<List<string>> comments(ExcelWorksheet worksheet, string evolution, string size, int row, string nature,
            List<string> commentList, string caption, List<int> auxiliar, string Coefficient , List<string> commentListSize,
            List<string> commentListCoefficient, List<double> peers, List<double> reference, int columnAux)
        {
            List<List<string>> output = new List<List<string>>();

            List<string> positiveComment = new List<string>();
            List<string> NegativeComment = new List<string>();
            List<string> OtherComment = new List<string>();

            string helper = null;

            List<int> color = new List<int>();
            if (caption == "Revenue" || caption == "Gross Margin" || caption == "Operating income Margin" || caption == "EBITDA Margin" || caption == "Net Income Margin")
            {
                helper = "PL";
            }
            else
            {
                helper = "BS";
            }

            double cagrRef = CAGR(reference);
            double cagrPeers = CAGR(peers);
            if (helper == "PL")
            {
                if (evolution == "++")
                {
                    //worksheet.Cells[row, 11].Style.Fill.PatternType = ExcelFillStyle.Solid;
                    //worksheet.Cells[row,11].Style.Fill.BackgroundColor.SetColor(Cores.DarkGreenWarning);
                    //worksheet.Cells[row, 12].Value = commentList[0];
                    //ExecSummary.ExecSummaryDetails.Add(new List<string> { caption, "DarkGreen", commentList[0] });
                    if (cagrRef > 0 && cagrPeers > 0)
                    {
                        positiveComment.Add(commentList[0]);
                        color.Add(1);
                        output.Add(new List<string> { "evolution", "increasing", "DarkGreen", caption });
                    }
                    else if (cagrRef > 0)
                    {
                        positiveComment.Add(commentList[7]);
                        color.Add(1);
                        output.Add(new List<string> { "evolution", "increasing", "DarkGreen", caption });
                    }
                    else
                    {
                        positiveComment.Add(commentList[5]);
                        color.Add(1);
                        output.Add(new List<string> { "evolution", "increasing", "DarkGreen", caption });
                    }


                }
                else if (evolution == "+")
                {

                    if (cagrRef > 0 && cagrPeers > 0)
                    {
                        positiveComment.Add(commentList[1]);
                        color.Add(2);
                        output.Add(new List<string> { "evolution", "increasing", "LightGreen", caption });
                    }
                    else if (cagrRef > 0)
                    {
                        positiveComment.Add(commentList[7]);
                        color.Add(2);
                        output.Add(new List<string> { "evolution", "increasing", "LightGreen", caption });
                    }
                    else
                    {
                        positiveComment.Add(commentList[7]);
                        color.Add(2);
                        output.Add(new List<string> { "evolution", "increasing", "LightGreen", caption });
                    }
                }
                else if (evolution == "+-")
                {
                    OtherComment.Add(commentList[2]);

                    output.Add(new List<string> { "evolution", "constant", "n.a.", caption });

                }
                else if (evolution == "-")
                {
                    if (cagrRef > 0 && cagrPeers > 0)
                    {
                        NegativeComment.Add(commentList[6]);
                        color.Add(4);
                        output.Add(new List<string> { "evolution", "decreasing", "Orange", caption });
                    }
                    else if (cagrPeers > 0)
                    {
                        NegativeComment.Add(commentList[8]);
                        color.Add(4);
                        output.Add(new List<string> { "evolution", "decreasing", "Orange", caption });
                    }
                    else
                    {
                        NegativeComment.Add(commentList[3]);
                        color.Add(4);
                        output.Add(new List<string> { "evolution", "decreasing", "Orange", caption });
                    }
                }
                else if (evolution == "--")
                {

                    if (cagrRef > 0 && cagrPeers > 0)
                    {
                        NegativeComment.Add(commentList[6]);
                        color.Add(5);
                        output.Add(new List<string> { "evolution", "decreasing", "Red", caption });
                    }
                    else if (cagrPeers > 0)
                    {
                        NegativeComment.Add(commentList[8]);
                        color.Add(5);
                        output.Add(new List<string> { "evolution", "decreasing", "Red", caption });
                    }
                    else
                    {
                        NegativeComment.Add(commentList[3]);
                        color.Add(5);
                        output.Add(new List<string> { "evolution", "decreasing", "Red", caption });
                    }
                }
                else
                {

                    color.Add(4);
                }




                if (caption != "Revenue")
                {

                    if (size == "++")
                    {
                        if (reference[reference.Count() - 1] > peers[peers.Count() - 1])
                        {
                            positiveComment.Add(commentListSize[0]);
                            color.Add(1);
                            output.Add(new List<string> { "size", "bigger", "DarkGreen", caption });
                        }

                    }
                    else if (size == "+")
                    {
                        if (reference[reference.Count() - 1] > peers[peers.Count() - 1])
                        {
                            positiveComment.Add(commentListSize[1]);
                            color.Add(2);
                            output.Add(new List<string> { "size", "bigger", "LightGreen", caption });
                        }
                    }
                    else if (size == "-")
                    {
                        if (reference[reference.Count() - 1] < peers[peers.Count() - 1])
                        {
                            NegativeComment.Add(commentListSize[3]);
                            color.Add(3);
                            output.Add(new List<string> { "size", "smaller", "Yellow", caption });
                        }
                    }
                    else if (size == "--")
                    {
                        if (reference[reference.Count() - 1] < peers[peers.Count() - 1])
                        {
                            NegativeComment.Add(commentListSize[4]);
                            color.Add(5);
                            output.Add(new List<string> { "size", "smaller", "Red", caption });
                        }
                    }


                }
                if (Coefficient == "++")
                {
                    NegativeComment.Add(commentListCoefficient[4]);
                    color.Add(4);
                    output.Add(new List<string> { "coeficient", "high", "Orange", caption });
                }
                else if (Coefficient == "+")
                {
                    NegativeComment.Add(commentListCoefficient[3]);
                    color.Add(4);
                    output.Add(new List<string> { "coeficient", "high", "Orange", caption });
                }
                else if (Coefficient == "+-")
                {
                    OtherComment.Add(commentListCoefficient[2]);

                    output.Add(new List<string> { "coeficient", "constant", "n.a.", caption });
                }
                else if (Coefficient == "-")
                {
                    positiveComment.Add(commentListCoefficient[1]);
                    color.Add(2);
                    output.Add(new List<string> { "coeficient", "low", "LighGreen", caption });
                }
                else if (Coefficient == "--")
                {
                    positiveComment.Add(commentListCoefficient[0]);
                    color.Add(1);
                    output.Add(new List<string> { "coeficient", "low", "DarkGreen", caption });
                }
                else
                {

                }
            }
            else
            {
                if (evolution == "++")
                {
                    OtherComment.Add(commentList[0]);
                    color.Add(3);
                }
                else if (evolution == "+")
                {
                    OtherComment.Add(commentList[0]);
                    color.Add(3);
                }
                else if (evolution == "-")
                {
                    OtherComment.Add(commentList[4]);
                    color.Add(3);
                }
                else if (evolution == "--")
                {
                    OtherComment.Add(commentList[4]);
                    color.Add(3);
                }

                if (size == "++")
                {
                    OtherComment.Add(commentListSize[0]);
                    color.Add(3);
                }
                else if (size == "--")
                {
                    OtherComment.Add(commentListSize[4]);
                    color.Add(3);
                }

                if (Coefficient == "++")
                {
                    OtherComment.Add(commentListCoefficient[4]);
                    color.Add(3);
                }
                else if (Coefficient == "--")
                {
                    OtherComment.Add(commentListCoefficient[0]);
                    color.Add(3);
                }
            }
            

            CommentAndColor commentAndColor = new CommentAndColor();
            commentAndColor.Comment(worksheet, row, columnAux, color, positiveComment, NegativeComment, OtherComment);

            return output;
        }

        public void ExecutiveSummaryFromIncomestatementAnalysis(List<List<string>> output)
        {
            List<double> evol = new List<double>();
            List<double> siz = new List<double>();
            List<double> coefic = new List<double>();
            if (output.Count()>0)
            {
                for (int i = 0; i < output.Count(); i++)
                {
                    if (output[i][0] == "evolution")
                    {
                        if (output[i][3] == "Revenue")
                        {
                            if (output[i][1] == "increasing")
                            {
                                ExecSummary.ExecSummaryDetails.Add( new List<string> { output[i][3], output[i][2], output[i][3] + " rising faster/decreasing at a lower pace than competitors. " });
                            }
                            else if (output[i][1] == "decreasing")
                            {
                                ExecSummary.ExecSummaryDetails.Add(new List<string> { output[i][3], output[i][2], output[i][3] + " decreasing faster/rising at a lower pace than competitors. " });
                            }
                        }
                        else if (output[i][3] == "Gross Margin")
                        {

                            if (output[i][1] == "increasing")
                            {
                                ExecSummary.ExecSummaryDetails.Add(new List<string> { output[i][3], output[i][2], output[i][3] + " rising faster/decreasing at a lower pace than competitors. " });
                                evol.Add(1);

                            }
                            else if (output[i][1] == "decreasing")
                            {
                                ExecSummary.ExecSummaryDetails.Add(new List<string> { output[i][3], output[i][2], output[i][3] + " decreasing faster/rising at a lower pace than competitors. " });
                                evol.Add(0);
                            }
                            else
                            {
                                evol.Add(0.5);
                            }

                        }
                        else
                        {
                            if (output[i][1] == "increasing")
                            {
                                if (evol.Count() == 0)
                                {
                                    ExecSummary.ExecSummaryDetails.Add(new List<string> { output[i][3], output[i][2], output[i][3] + " rising faster/decreasing at a lower pace than competitors. " });
                                    evol.Add(1);
                                }
                                else
                                {
                                    if (evol.Max() < 1)
                                    {
                                        ExecSummary.ExecSummaryDetails.Add(new List<string> { output[i][3], output[i][2], output[i][3] + " rising faster/decreasing at a lower pace than competitors. " });
                                        evol.Add(1);
                                    }
                                }


                            }
                            else if (output[i][1] == "decreasing")
                            {
                                if (evol.Count()== 0)
                                {
                                    ExecSummary.ExecSummaryDetails.Add(new List<string> { output[i][3], output[i][2], output[i][3] + " decreasing faster/rising at a lower pace than competitors. " });
                                    evol.Add(1);
                                }
                                else
                                {
                                    if (evol.Min() > 0 || evol == null)
                                    {
                                        ExecSummary.ExecSummaryDetails.Add(new List<string> { output[i][3], output[i][2], output[i][3] + " decreasing faster/rising at a lower pace than competitors. " });
                                        evol.Add(0);
                                    }
                                }

                            }
                        }
                    }
                    else if (output[i][0] == "size")
                    {
                        if (output[i][1] == "bigger")
                        {
                            if (siz.Count()==0)
                            {
                                ExecSummary.ExecSummaryDetails.Add(new List<string> { output[i][3], output[i][2], output[i][3] + " is higher than those of its competitors. " });
                                siz.Add(1);
                            }
                            else
                            {
                                if (siz.Max() < 1 || siz == null)
                                {
                                    ExecSummary.ExecSummaryDetails.Add(new List<string> { output[i][3], output[i][2], output[i][3] + " is higher than those of its competitors. " });
                                    siz.Add(1);
                                }
                            }
                        }
                        else if (output[i][1] == "smaller")
                        {
                            if (siz.Count()==0)
                            {
                                ExecSummary.ExecSummaryDetails.Add(new List<string> { output[i][3], output[i][2], output[i][3] + " is lower than those of its competitors.. " });
                                siz.Add(0);
                            }
                            else
                            {
                                if (siz.Min() > 0 || siz == null)
                                {
                                    ExecSummary.ExecSummaryDetails.Add(new List<string> { output[i][3], output[i][2], output[i][3] + " is lower than those of its competitors.. " });
                                    siz.Add(0);
                                }
                            }


                        }
                        else
                        {
                            siz.Add(0.5);
                        }
                    }
                    else
                    {

                        if (output[i][1] == "high")
                        {
                            if (coefic.Count()==0)
                            {
                                ExecSummary.ExecSummaryDetails.Add(new List<string> { output[i][3], output[i][2], output[i][3] + " volatility is higher than those of its competitors. " });
                                
                                if (output[i][3] != "Revenue")
                                {
                                    coefic.Add(1);
                                }
                            }
                            else
                            {
                                if (coefic.Max() < 1)
                                {
                                    ExecSummary.ExecSummaryDetails.Add(new List<string> { output[i][3], output[i][2], output[i][3] + " volatility is higher than those of its competitors. " });

                                    if (output[i][3] != "Revenue")
                                    {
                                        coefic.Add(1);
                                    }
                                }
                            }
                        }
                        else if (output[i][1] == "low")
                        {
                            if (coefic.Count() == 0)
                            {
                                ExecSummary.ExecSummaryDetails.Add(new List<string> { output[i][3], output[i][2], output[i][3] + " volatility is lower than those of its competitors.. " });
                                if (output[i][3] != "Revenue")
                                {
                                    coefic.Add(0);
                                }
                            }
                            else
                            {
                                if (coefic.Min() > 0 || coefic == null)
                                {
                                    ExecSummary.ExecSummaryDetails.Add(new List<string> { output[i][3], output[i][2], output[i][3] + " volatility is lower than those of its competitors.. " });
                                    if (output[i][3] != "Revenue")
                                    {
                                        coefic.Add(0);
                                    }
                                }
                            }


                        }
                        else
                        {
                            coefic.Add(0.5);
                        }
                    }

                }
            }


        }
    }
}
