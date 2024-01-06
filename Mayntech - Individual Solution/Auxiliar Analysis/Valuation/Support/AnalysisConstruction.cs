using Mayntech___Individual_Solution.Auxiliar.Auxiliar.Other;
using Mayntech___Individual_Solution.Auxiliar.Valuation.Support;
using Mayntech___Individual_Solution.Pages;
using OfficeOpenXml;
using OfficeOpenXml.Sparkline;
using OfficeOpenXml.Style;
using System.Collections.Generic;
using System.Drawing;
using System.Linq;
using System.Runtime.ConstrainedExecution;

namespace Mayntech___Individual_Solution.Auxiliar.Analysis
{
    public class AnalysisConstruction : ExcelNextCol
    {

        public void CommonAnalysisRefRevenue(ExcelWorksheet worksheet, int row, int NumberOfYears, string CaptionName, List<double> PropertyValues, List<double> revenueValues, string Sinal)
        {

            //Escreve as primeiras linhas com os valores reportados, crescimento anual e percentagem da referência
            int col = 2;
            worksheet.Cells[row, col].Value = CaptionName;
            worksheet.Cells[row, col].Style.Font.Bold = true;
            worksheet.Cells[row + 1, col].Value = "As reported";
            worksheet.Cells[row + 2, col].Value = "% change";
            worksheet.Cells[row + 2, col].Style.Font.Color.SetColor(Color.Gray);

            for (int i = 0; i < 4; i++)
            {
                worksheet.Cells[row + 1 + i, col].Style.Fill.PatternType = ExcelFillStyle.Solid;
                worksheet.Cells[row + 1 + i, col].Style.Fill.BackgroundColor.SetColor(Cores.CorBackgroundAnalysis);
            }

            worksheet.Cells[row + 4, col].Value = "as % of revenue";
            worksheet.Cells[row + 4, col].Style.Font.Italic = true;

            for (int i = 0; i < NumberOfYears; i++)
            {
                worksheet.Row(row - 1).Height = 4;
                int colunaAux = 2 + i;
                string aux = '"' + CaptionName + '"';
                worksheet.Cells[row + 1, col + i + 1].Formula = "=VLOOKUP(" + aux + ",'P&L'!B:AZ," + colunaAux + ",0)";
                worksheet.Cells[row + 1, col + i + 1].Style.Numberformat.Format = "#,##0;(#,##0);-";
                worksheet.Cells[row + 1, col + i + 1].Style.Fill.PatternType = ExcelFillStyle.Solid;
                worksheet.Cells[row + 1, col + i + 1].Style.Fill.BackgroundColor.SetColor(Cores.CorBackgroundAnalysis);

                worksheet.Cells[row + 2, col + i + 1].Style.Fill.PatternType = ExcelFillStyle.Solid;
                worksheet.Cells[row + 2, col + i + 1].Style.Fill.BackgroundColor.SetColor(Cores.CorBackgroundAnalysis);

                worksheet.Cells[row + 3, col + i + 1].Style.Fill.PatternType = ExcelFillStyle.Solid;
                worksheet.Cells[row + 3, col + i + 1].Style.Fill.BackgroundColor.SetColor(Cores.CorBackgroundAnalysis);

                worksheet.Cells[row + 4, col + i + 1].Style.Fill.PatternType = ExcelFillStyle.Solid;
                worksheet.Cells[row + 4, col + i + 1].Style.Fill.BackgroundColor.SetColor(Cores.CorBackgroundAnalysis);


                //% de crescimento anual
                if (i != 0)
                {
                    string ColAnoAnterior = GetExcelColumnName(col + i);
                    string ColAno = GetExcelColumnName(col + 1 + i);
                    int rowAux = row + 1;
                    worksheet.Cells[row + 2, col + i + 1].Formula = "=" + ColAno + rowAux + "/" + ColAnoAnterior + rowAux;
                    worksheet.Cells[row + 2, col + i + 1].Style.Numberformat.Format = "#,##0;(#,##0);-";
                    worksheet.Cells[row + 2, col + i + 1].Style.Numberformat.Format = "0.00%";
                    worksheet.Cells[row + 2, col + i + 1].Style.Font.Color.SetColor(Color.Gray);
                }


                //A divisão com a referencia
                string ColAux = GetExcelColumnName(col + 1 + i);
                int rowAux1 = row + 1;
                if (Sinal == "Negative")
                {
                    worksheet.Cells[row + 4, col + i + 1].Formula = "=-" + ColAux + rowAux1 + "/ 'P&L'!" + ColAux + "5";
                }
                else if (Sinal != "Negative")
                {
                    worksheet.Cells[row + 4, col + i + 1].Formula = "=" + ColAux + rowAux1 + "/ 'P&L'!" + ColAux + "5";
                }

                worksheet.Cells[row + 4, col + i + 1].Style.Numberformat.Format = "0.00%";
                worksheet.Cells[row + 4, col + i + 1].Style.Font.Italic = true;

            }
            string ColAux1 = GetExcelColumnName(col + 1 + NumberOfYears);
            string ColAux2 = GetExcelColumnName(col  + NumberOfYears);
            int rowAuxiliar = row + 4;
            var sparklineLine = worksheet.SparklineGroups.Add(eSparklineType.Line, worksheet.Cells[ColAux1 + rowAuxiliar], worksheet.Cells["C" + rowAuxiliar + ":" + ColAux2 + rowAuxiliar]);



            //começa a construção dos comentários

            worksheet.Cells[row, col + 2 + NumberOfYears].Value = "Confidence Level";
            worksheet.Cells[row, col + 2 + NumberOfYears].Style.Font.Bold = true;
            worksheet.Cells[row, col + 2 + NumberOfYears].Style.Border.Bottom.Style = ExcelBorderStyle.Thin;
            worksheet.Columns[col + 2 + NumberOfYears].Width = 16;

            worksheet.Cells[row, col + 3 + NumberOfYears].Value = "Estimated Interval";
            worksheet.Cells[row, col + 3 + NumberOfYears].Style.Font.Bold = true;
            worksheet.Cells[row, col + 3 + NumberOfYears].Style.Border.Bottom.Style = ExcelBorderStyle.Thin;
            worksheet.Columns[col + 3 + NumberOfYears].Width = 20;


            worksheet.Cells[row, col + 4 + NumberOfYears].Value = "Estimated value";
            worksheet.Cells[row, col + 4 + NumberOfYears].Style.Font.Bold = true;
            worksheet.Cells[row, col + 4 + NumberOfYears].Style.Border.Bottom.Style = ExcelBorderStyle.Thin;
            worksheet.Columns[col + 4 + NumberOfYears].Width = 15;

            worksheet.Cells[row, col + 5 + NumberOfYears].Value = "User's Comments";
            worksheet.Cells[row, col + 5 + NumberOfYears].Style.Font.Bold = true;
            worksheet.Cells[row, col + 5 + NumberOfYears].Style.Border.Bottom.Style = ExcelBorderStyle.Thin;
            worksheet.Columns[col + 5 + NumberOfYears].Width = 15;

            string startColumn = GetExcelColumnName(col + 3 + NumberOfYears);
            int startRow = row + 1;
            int endRow = row + 4;
            worksheet.Cells[startColumn + startRow + ":" + startColumn + endRow].Merge = true;
            worksheet.Cells[row + 1, col + 3 + NumberOfYears].Style.VerticalAlignment = ExcelVerticalAlignment.Top; ;
            worksheet.Cells[row + 1, col + 3 + NumberOfYears].Style.WrapText = true;

            //remove os anos que não estão a ser usados
            int anosARemover = PropertyValues.Count() - NumberOfYears;

            PropertyValues.RemoveRange(0, anosARemover);
            revenueValues.RemoveRange(0, anosARemover);
            double[] values = new double[PropertyValues.Count()];
            List<double> Values = new List<double>();
            List<double> xvalues = new List<double>();

            //Comparação com referencia
            for (int a = 0; a < PropertyValues.Count(); a++)
            {
                double aux = PropertyValues[a] / revenueValues[a];

                values[a] = aux;
                Values.Add(aux);
                xvalues.Add(a);
            }


            //Calculate the standard deviation
            double average = values.Average();
            double sumOfSquaresOfDifferences = values.Select(val => (val - average) * (val - average)).Sum();
            double sd = Math.Sqrt(sumOfSquaresOfDifferences / values.Length);



            IDictionary<string, string> comments = new Dictionary<string, string>();
            AnalysisSupport support = new AnalysisSupport();

            support.Estimations(worksheet, row, col + 2 + NumberOfYears, revenueValues, Values);

        }

        public void Output(List<string> Listwarning, ExcelWorksheet worksheet, int row, int col, int NumberOfYears)
        {
            string warning = null;
            if (Listwarning == null)
            {
                return;
            }
            try
            {
                if (Listwarning.Contains("Red"))
                {
                    warning = "Red";
                }
                else
                {
                    if (Listwarning.Contains("Orange"))
                    {
                        warning = "Orange";
                    }
                    else
                    {
                        if (Listwarning.Contains("Yellow"))
                        {
                            warning = "Yellow";
                        }
                        else
                        {
                            if (Listwarning.Contains("Green"))
                            {
                                warning = "Green";
                            }
                            else
                            {
                                warning = "Gray";
                            }
                        }
                    }
                }
            }
            finally
            {
                worksheet.Cells[row + 1, col + 2 + NumberOfYears].Style.Fill.PatternType = ExcelFillStyle.Solid;
                if (warning == "Red")
                {
                    worksheet.Cells[row + 1, col + 2 + NumberOfYears].Style.Fill.BackgroundColor.SetColor(Color.Red);
                }
                if (warning == "Yellow")
                {
                    worksheet.Cells[row + 1, col + 2 + NumberOfYears].Style.Fill.BackgroundColor.SetColor(Color.Yellow);
                }
                if (warning == "Green")
                {
                    worksheet.Cells[row + 1, col + 2 + NumberOfYears].Style.Fill.BackgroundColor.SetColor(Color.Green);
                }
                if (warning == "Gray")
                {
                    worksheet.Cells[row + 1, col + 2 + NumberOfYears].Style.Fill.BackgroundColor.SetColor(Color.Gray);
                }
                if (warning == "Orange")
                {
                    worksheet.Cells[row + 1, col + 2 + NumberOfYears].Style.Fill.BackgroundColor.SetColor(Color.Orange);
                }
            }


        }


        public void CommonAnalysisRefRevenueBS(ExcelWorksheet worksheet, int row, int NumberOfYears, string CaptionName, List<double> PropertyValues, List<double> revenueValues, string Sinal)
        {

            //Escreve as primeiras linhas com os valores reportados, crescimento anual e percentagem da referência
            int col = 2;
            worksheet.Cells[row, col].Value = CaptionName;
            worksheet.Cells[row, col].Style.Font.Bold = true;
            worksheet.Cells[row + 1, col].Value = "As reported";
            worksheet.Cells[row + 2, col].Value = "% change";
            worksheet.Cells[row + 2, col].Style.Font.Color.SetColor(Color.Gray);

            for (int i = 0; i < 4; i++)
            {
                worksheet.Cells[row + 1 + i, col].Style.Fill.PatternType = ExcelFillStyle.Solid;
                worksheet.Cells[row + 1 + i, col].Style.Fill.BackgroundColor.SetColor(Cores.CorBackgroundAnalysis);
            }

            worksheet.Cells[row + 4, col].Value = "as % of revenue";
            worksheet.Cells[row + 4, col].Style.Font.Italic = true;

            for (int i = 0; i < NumberOfYears; i++)
            {
                worksheet.Row(row - 1).Height = 4;
                int colunaAux = 2 + i;
                string aux = '"' + CaptionName + '"';
                worksheet.Cells[row + 1, col + i + 1].Formula = "=VLOOKUP(" + aux + ",'BS'!B:AZ," + colunaAux + ",0)";
                worksheet.Cells[row + 1, col + i + 1].Style.Numberformat.Format = "#,##0;(#,##0);-";
                worksheet.Cells[row + 1, col + i + 1].Style.Fill.PatternType = ExcelFillStyle.Solid;
                worksheet.Cells[row + 1, col + i + 1].Style.Fill.BackgroundColor.SetColor(Cores.CorBackgroundAnalysis);

                worksheet.Cells[row + 2, col + i + 1].Style.Fill.PatternType = ExcelFillStyle.Solid;
                worksheet.Cells[row + 2, col + i + 1].Style.Fill.BackgroundColor.SetColor(Cores.CorBackgroundAnalysis);

                worksheet.Cells[row + 3, col + i + 1].Style.Fill.PatternType = ExcelFillStyle.Solid;
                worksheet.Cells[row + 3, col + i + 1].Style.Fill.BackgroundColor.SetColor(Cores.CorBackgroundAnalysis);

                worksheet.Cells[row + 4, col + i + 1].Style.Fill.PatternType = ExcelFillStyle.Solid;
                worksheet.Cells[row + 4, col + i + 1].Style.Fill.BackgroundColor.SetColor(Cores.CorBackgroundAnalysis);


                //% de crescimento anual
                if (i != 0)
                {
                    string ColAnoAnterior = GetExcelColumnName(col + i);
                    string ColAno = GetExcelColumnName(col + 1 + i);
                    int rowAux = row + 1;
                    worksheet.Cells[row + 2, col + i + 1].Formula = "=" + ColAno + rowAux + "/" + ColAnoAnterior + rowAux;
                    worksheet.Cells[row + 2, col + i + 1].Style.Numberformat.Format = "#,##0;(#,##0);-";
                    worksheet.Cells[row + 2, col + i + 1].Style.Numberformat.Format = "0.00%";
                    worksheet.Cells[row + 2, col + i + 1].Style.Font.Color.SetColor(Color.Gray);
                }


                //A divisão com a referencia
                string ColAux = GetExcelColumnName(col + 1 + i);
                int rowAux1 = row + 1;
                if (Sinal == "Negative")
                {
                    worksheet.Cells[row + 4, col + i + 1].Formula = "=-" + ColAux + rowAux1 + "/ 'P&L'!" + ColAux + "5";
                }
                else if (Sinal != "Negative")
                {
                    worksheet.Cells[row + 4, col + i + 1].Formula = "=" + ColAux + rowAux1 + "/ 'P&L'!" + ColAux + "5";
                }

                worksheet.Cells[row + 4, col + i + 1].Style.Numberformat.Format = "0.00%";
                worksheet.Cells[row + 4, col + i + 1].Style.Font.Italic = true;

            }
            string ColAux1 = GetExcelColumnName(col + 1 + NumberOfYears);
            string ColAux2 = GetExcelColumnName(col + NumberOfYears);
            int rowAuxiliar = row + 4;
            var sparklineLine = worksheet.SparklineGroups.Add(eSparklineType.Line, worksheet.Cells[ColAux1 + rowAuxiliar], worksheet.Cells["C" + rowAuxiliar + ":" + ColAux2 + rowAuxiliar]);



            //começa a construção dos comentários

            worksheet.Cells[row, col + 2 + NumberOfYears].Value = "Confidence Level";
            worksheet.Cells[row, col + 2 + NumberOfYears].Style.Font.Bold = true;
            worksheet.Cells[row, col + 2 + NumberOfYears].Style.Border.Bottom.Style = ExcelBorderStyle.Thin;
            worksheet.Columns[col + 2 + NumberOfYears].Width = 16;

            worksheet.Cells[row, col + 3 + NumberOfYears].Value = "Estimated Interval";
            worksheet.Cells[row, col + 3 + NumberOfYears].Style.Font.Bold = true;
            worksheet.Cells[row, col + 3 + NumberOfYears].Style.Border.Bottom.Style = ExcelBorderStyle.Thin;
            worksheet.Columns[col + 3 + NumberOfYears].Width = 20;


            worksheet.Cells[row, col + 4 + NumberOfYears].Value = "Estimated value";
            worksheet.Cells[row, col + 4 + NumberOfYears].Style.Font.Bold = true;
            worksheet.Cells[row, col + 4 + NumberOfYears].Style.Border.Bottom.Style = ExcelBorderStyle.Thin;
            worksheet.Columns[col + 4 + NumberOfYears].Width = 15;

            worksheet.Cells[row, col + 5 + NumberOfYears].Value = "User's Comments";
            worksheet.Cells[row, col + 5 + NumberOfYears].Style.Font.Bold = true;
            worksheet.Cells[row, col + 5 + NumberOfYears].Style.Border.Bottom.Style = ExcelBorderStyle.Thin;
            worksheet.Columns[col + 5 + NumberOfYears].Width = 15;

            string startColumn = GetExcelColumnName(col + 3 + NumberOfYears);
            int startRow = row + 1;
            int endRow = row + 4;
            worksheet.Cells[startColumn + startRow + ":" + startColumn + endRow].Merge = true;
            worksheet.Cells[row + 1, col + 3 + NumberOfYears].Style.VerticalAlignment = ExcelVerticalAlignment.Top; ;
            worksheet.Cells[row + 1, col + 3 + NumberOfYears].Style.WrapText = true;

            //remove os anos que não estão a ser usados
            int anosARemover = 0;
            int anosARemoverRevenue = 0;


            if (PropertyValues.Count()< NumberOfYears || revenueValues.Count()< NumberOfYears)
            {
                int aux = Math.Min(PropertyValues.Count(), revenueValues.Count());
                anosARemover = PropertyValues.Count() - aux;
                anosARemoverRevenue = revenueValues.Count() - aux;
            }
            else if (PropertyValues.Count()!=revenueValues.Count())
            {
                anosARemover = PropertyValues.Count() - NumberOfYears;
                anosARemoverRevenue = revenueValues.Count() - NumberOfYears;
            }
            else
            {
                anosARemover = PropertyValues.Count() - NumberOfYears;
                anosARemoverRevenue = anosARemover;
            }

            PropertyValues.RemoveRange(0, anosARemover);
            revenueValues.RemoveRange(0, anosARemoverRevenue);
            double[] values = new double[PropertyValues.Count()];
            List<double> Values = new List<double>();
            List<double> xvalues = new List<double>();

            //Comparação com referencia
            for (int a = 0; a < PropertyValues.Count(); a++)
            {
                double aux = PropertyValues[a] / revenueValues[a];

                values[a] = aux;
                Values.Add(aux);
                xvalues.Add(a);
            }


            //Calculate the standard deviation
            double average = values.Average();
            double sumOfSquaresOfDifferences = values.Select(val => (val - average) * (val - average)).Sum();
            double sd = Math.Sqrt(sumOfSquaresOfDifferences / values.Length);

            IDictionary<string, string> comments = new Dictionary<string, string>();
            AnalysisSupport support = new AnalysisSupport();

            support.Estimations(worksheet, row, col + 2 + NumberOfYears, revenueValues, Values);


        }

    }

}
