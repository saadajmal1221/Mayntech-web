using OfficeOpenXml.Sparkline;
using OfficeOpenXml.Style;
using OfficeOpenXml;
using System.Drawing;
using Mayntech___Individual_Solution.Auxiliar.Auxiliar.Other;

namespace Mayntech___Individual_Solution.Auxiliar.Analysis.Analysis_Support
{
    public class RevenueAnalysis : ExcelNextCol
    {
        public void RevenueAnalysisConstruction(ExcelWorksheet worksheet, int row, int NumberOfYears, List<double> revenueValues)
        {
            int col = 2;
            //Escreve as primeiras linhas com os valores reportados, crescimento anual e percentagem da referência
            try
            {

                worksheet.Cells[row, col].Value = "Revenue";
                worksheet.Cells[row, col].Style.Font.Bold = true;
                worksheet.Cells[row + 1, col].Value = "As reported";
                worksheet.Cells[row + 2, col].Value = "% change";
                worksheet.Cells[row + 2, col].Style.Font.Color.SetColor(Color.Gray);
            }
            catch (Exception)
            {

                throw;
            }


            for (int i = 0; i < 2; i++)
            {
                worksheet.Cells[row + 1 + i, col].Style.Fill.PatternType = ExcelFillStyle.Solid;
                worksheet.Cells[row + 1 + i, col].Style.Fill.BackgroundColor.SetColor(Cores.CorBackgroundAnalysisRevenue);
            }



            for (int i = 0; i < NumberOfYears; i++)
            {
                worksheet.Row(row - 1).Height = 4;
                int colunaAux = 2 + i;
                worksheet.Cells[row + 1, col + i + 1].Formula = "=VLOOKUP(" +'"'+ "Revenue" +'"'+ ",'P&L'!B:AZ," + colunaAux + ",0)";
                worksheet.Cells[row + 1, col + i + 1].Style.Numberformat.Format = "#,##0;(#,##0);-";
                worksheet.Cells[row + 1, col + i + 1].Style.Fill.PatternType = ExcelFillStyle.Solid;
                worksheet.Cells[row + 1, col + i + 1].Style.Fill.BackgroundColor.SetColor(Cores.CorBackgroundAnalysisRevenue);

                worksheet.Cells[row + 2, col + i + 1].Style.Fill.PatternType = ExcelFillStyle.Solid;
                worksheet.Cells[row + 2, col + i + 1].Style.Fill.BackgroundColor.SetColor(Cores.CorBackgroundAnalysisRevenue);


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


            }
            string ColAux1 = GetExcelColumnName(col + 1 + NumberOfYears);
            string ColAux2 = GetExcelColumnName(col + NumberOfYears);
            int rowAuxiliar = row + +1;
            var sparklineLine = worksheet.SparklineGroups.Add(eSparklineType.Line, worksheet.Cells[ColAux1 + rowAuxiliar], worksheet.Cells["C" + rowAuxiliar + ":" + ColAux2 + rowAuxiliar]);



            //começa a construção dos comentários

            worksheet.Cells[row, col + 2 + NumberOfYears].Value = "Importance";
            worksheet.Cells[row, col + 2 + NumberOfYears].Style.Font.Bold = true;
            worksheet.Cells[row, col + 2 + NumberOfYears].Style.Border.Bottom.Style = ExcelBorderStyle.Thin;
            worksheet.Columns[col + 2 + NumberOfYears].Width = 12;

            worksheet.Cells[row, col + 3 + NumberOfYears].Value = "Comments";
            worksheet.Cells[row, col + 3 + NumberOfYears].Style.Font.Bold = true;
            worksheet.Cells[row, col + 3 + NumberOfYears].Style.Border.Bottom.Style = ExcelBorderStyle.Thin;
            worksheet.Columns[col + 3 + NumberOfYears].Width = 40;


            worksheet.Cells[row, col + 4 + NumberOfYears].Value = "User's Comments";
            worksheet.Cells[row, col + 4 + NumberOfYears].Style.Font.Bold = true;
            worksheet.Cells[row, col + 4 + NumberOfYears].Style.Border.Bottom.Style = ExcelBorderStyle.Thin;
            worksheet.Columns[col + 4 + NumberOfYears].Width = 25;
            string startColumn = GetExcelColumnName(col + 3 + NumberOfYears);
            int startRow = row + 1;
            int endRow = row + 2;
            worksheet.Cells[startColumn + startRow + ":" + startColumn + endRow].Merge = true;
            worksheet.Cells[row + 1, col + 3 + NumberOfYears].Style.VerticalAlignment = ExcelVerticalAlignment.Top; ;
            worksheet.Cells[row + 1, col + 3 + NumberOfYears].Style.WrapText = true;

            //remove os anos que não estão a ser usados
            int anosARemover = revenueValues.Count() - NumberOfYears;


            revenueValues.RemoveRange(0, anosARemover);

            List<double> xvalues = new List<double>();

            //Comparação com referencia
            for (int a = 0; a < revenueValues.Count(); a++)
            {

                xvalues.Add(a);
            }


            //Calculate the standard deviation
            double average = revenueValues.Average();
            double sumOfSquaresOfDifferences = revenueValues.Select(val => (val - average) * (val - average)).Sum();
            double sd = Math.Sqrt(sumOfSquaresOfDifferences / revenueValues.Count());

            LinearRegression regression = new LinearRegression();
            List<double> outputLinearRegression = regression.LinearRegressionCalculation(xvalues, revenueValues);
            double teste = outputLinearRegression[0] / average;

            IDictionary<string, string> comments = new Dictionary<string, string>();

            comments.Add("Size", "N/A");
            comments.Add("Nature", "N/A");

            if (outputLinearRegression[2] < 0.4)
            {
                if (sd < 0.1)
                {
                    comments.Add("Sd", "Small");
                }
                else if (sd >= 0.1 && sd < 0.2)
                {
                    comments.Add("Sd", "Medium");
                }
                else
                {
                    comments.Add("Sd", "Large");
                }
            }
            else if (outputLinearRegression[2] < 0.6 && outputLinearRegression[2] >= 0.4)
            {

                if (sd >= 0.3 && sd< 0.6)
                {
                    comments.Add("Sd", "Medium");
                }
                else if (sd >= 0.6)
                {
                    comments.Add("Sd", "Large");
                }

            }
            else
            {
                comments.Add("Sd", "N/A");
            }


            //Comentário do slope
            if (outputLinearRegression[0]/average < -0.004 && outputLinearRegression[2] > 0.4)
            {
                comments.Add("Slope", "Large Negative");
            }
            if (outputLinearRegression[0]/average < -0.004 && outputLinearRegression[2] <= 0.4 && outputLinearRegression[2] >= 0.2)
            {
                comments.Add("Slope", "Medium Negative");
            }
            if (outputLinearRegression[0]/average < -0.004 && outputLinearRegression[2] < 0.2)
            {
                comments.Add("Slope", "N/A");
            }
            if (outputLinearRegression[0]/average < -0.0025 && outputLinearRegression[0]/average >= -0.004)
            {
                comments.Add("Slope", "Medium Negative");
            }
            if (outputLinearRegression[0]/average < -0.0015 && outputLinearRegression[0]/average >= -0.0025)
            {
                comments.Add("Slope", "Small Negative");
            }
            if (outputLinearRegression[0]/average >= -0.0015 && outputLinearRegression[0]/average < 0.000)
            {
                comments.Add("Slope", "N/A");
            }
            if (outputLinearRegression[0]/average >= 0.000 && outputLinearRegression[0]/average < 0.0015)
            {
                comments.Add("Slope", "N/A");
            }
            if (outputLinearRegression[0]/average >= 0.0015 && outputLinearRegression[0]/average < 0.0025)
            {
                comments.Add("Slope", "Small Positive");
            }
            if (outputLinearRegression[0]/average >= 0.0025 && outputLinearRegression[0]/average < 0.004)
            {
                comments.Add("Slope", "Medium Positive");
            }
            if (outputLinearRegression[0]/average >= 0.004 && outputLinearRegression[2] > 0.4)
            {
                comments.Add("Slope", "Large Positive");
            }
            if (outputLinearRegression[0]/average >= 0.004 && outputLinearRegression[2] <= 0.4)
            {
                comments.Add("Slope", "Medium Positive");
            }



        }

        
    }
}
