using Mayntech___Individual_Solution.Auxiliar.Auxiliar.Other;
using OfficeOpenXml;
using OfficeOpenXml.Style;
using System.Drawing;
using System.Security.Cryptography.X509Certificates;

namespace Mayntech___Individual_Solution.Auxiliar.Executive_Summary
{
    public class ExecSummaryConstruction
    {
        public void ConstructionExecSummary(ExcelPackage package, string CompanyName)
        {
            var worksheet = package.Workbook.Worksheets.Add("Executive Summary");
            worksheet.View.ShowGridLines = false;
            worksheet.View.ZoomScale = 80;
            int row = 2;
            int col = 2;

            worksheet.Cells[row, col].Value = "Executive Summary - " + CompanyName;
            worksheet.Cells[row, col].Style.Font.Color.SetColor(Color.White);
            worksheet.Cells[row, col].Style.Font.Bold = true;

            for (int i = 0; i < 3; i++)
            {
                worksheet.Cells[row, col + i].Style.Fill.PatternType = ExcelFillStyle.Solid;
                worksheet.Cells[row, col + i].Style.Fill.BackgroundColor.SetColor(Cores.CorPrincipal);
            }





            List<List<string>> positive = new List<List<string>>();
            List<List<string>> negative = new List<List<string>>();
            List<int> aux = new List<int>();
            for (int i = 0; i < ExecSummary.ExecSummaryDetails.Count(); i++)
            {
                if (ExecSummary.ExecSummaryDetails[i][1] == "DarkGreen" || ExecSummary.ExecSummaryDetails[i][1] == "LightGreen" && ExecSummary.ExecSummaryDetails[i][0]!="Cash" && ExecSummary.ExecSummaryDetails[i][0] != "CashFromOperations" && ExecSummary.ExecSummaryDetails[i][0] != "Goodwill")
                {
                    positive.Add(ExecSummary.ExecSummaryDetails[i]);
                }
                else if (ExecSummary.ExecSummaryDetails[i][1] == "Yellow" || ExecSummary.ExecSummaryDetails[i][1] == "Orange" || ExecSummary.ExecSummaryDetails[i][1] == "Red" && ExecSummary.ExecSummaryDetails[i][0] != "Cash" && ExecSummary.ExecSummaryDetails[i][0] != "CashFromOperations" && ExecSummary.ExecSummaryDetails[i][0] != "Goodwill")
                {
                    negative.Add(ExecSummary.ExecSummaryDetails[i]);
                }

            }

            worksheet.Cells[row+2, 2].Value = "Negative Findings";
            worksheet.Cells[row+2, 2].Style.Font.Bold = true;

            ExecAuxiliar(negative, worksheet, row + 3);
            row += negative.Count() + 5;

            worksheet.Cells[row+1, 2].Value = "Positive Findings";
            worksheet.Cells[row+1, 2].Style.Font.Bold = true;
            ExecAuxiliar(positive, worksheet, row + 2);
            row += positive.Count() + 5;

            //LegendConstructor(worksheet, row);

            ExecSummary.ExecSummaryDetails.Clear();
        }
        public void LegendConstructor(ExcelWorksheet worksheet, int row)
        {
            worksheet.Cells[row, 2].Value = "Legend:";
            worksheet.Cells[row, 2].Style.Font.Bold = true;

            worksheet.Cells[row+1, 3].Value = "High importance - Positive Finding";
            worksheet.Cells[row +1, 2].Style.Fill.PatternType = ExcelFillStyle.Solid;
            worksheet.Cells[row + 1, 2].Style.Fill.BackgroundColor.SetColor(Cores.DarkGreenWarning);


            worksheet.Cells[row + 2, 3].Value = "Medium importance - Positive Finding";
            worksheet.Cells[row + 2, 2].Style.Fill.PatternType = ExcelFillStyle.Solid;
            worksheet.Cells[row + 2, 2].Style.Fill.BackgroundColor.SetColor(Cores.LightGreenWarning);

            worksheet.Cells[row + 3, 3].Value = "High importance - Negative Finding";
            worksheet.Cells[row + 3, 2].Style.Fill.PatternType = ExcelFillStyle.Solid;
            worksheet.Cells[row + 3, 2].Style.Fill.BackgroundColor.SetColor(Cores.RedWarning);

            worksheet.Cells[row + 4, 3].Value = "Medium importance - Negative Finding";
            worksheet.Cells[row + 4, 2].Style.Fill.PatternType = ExcelFillStyle.Solid;
            worksheet.Cells[row + 4, 2].Style.Fill.BackgroundColor.SetColor(Cores.OrangeWarning);

            worksheet.Cells[row + 5, 3].Value = "Low importance - Negative Finding";
            worksheet.Cells[row + 5, 2].Style.Fill.PatternType = ExcelFillStyle.Solid;
            worksheet.Cells[row + 5, 2].Style.Fill.BackgroundColor.SetColor(Cores.YellowWarning);
        }
        public void ExecAuxiliar(List<List<string>> Comments, ExcelWorksheet worksheet, int row)
        {

            int col = 2;

            for (int i = 0; i < 3; i++)
            {



                worksheet.Cells[row, col + i].Style.Fill.PatternType = ExcelFillStyle.Solid;
                worksheet.Cells[row, col + i].Style.Fill.BackgroundColor.SetColor(Cores.CorCinzentoBackground);


                worksheet.Cells[row, col].Value = "Indicator";
                worksheet.Cells[row, col].Style.Font.Bold = true;
                worksheet.Column(col).Width = 12;

                worksheet.Cells[row, col + 1].Value = "Comments";
                worksheet.Cells[row, col + 1].Style.Font.Bold = true;
                worksheet.Column(col + 1).Width = 100;

                worksheet.Cells[row, col + 2].Value = "User's Comments";
                worksheet.Cells[row, col + 2].Style.Font.Bold = true;
                worksheet.Column(col + 2).Width = 20;

                worksheet.Rows[row + 1].Height = 2;
                int counter = 0;

                for (int a = 0; a < Comments.Count(); a++)
                {
                    if (Comments[a][0] != "Cash" && Comments[a][0] != "Goodwill" && Comments[a][0] != "CashFromOperations" &&
                       Comments[a][0] != "Quick Ratio" && Comments[a][0] != "Cash ratio" && Comments[a][0] != "Days of Inventory Outstanding (DIO)" &&
                       Comments[a][0] != "Days of Sales Outstanding (DSO)" && Comments[a][0] != "Days of Payables Outstanding (DPO)")
                    {                        
                        indicatorTranslator(Comments[a], worksheet, row + counter);

                        worksheet.Cells[row + 2 + counter, 3].Value = Comments[a][2];
                        counter++;
                    }


                }

            }

        }

        public void indicatorTranslator(List<string> input, ExcelWorksheet worksheet, int row)
        {
            if (input[1] =="DarkGreen")
            {
                worksheet.Cells[row+2, 2].Value = "High";
                //worksheet.Cells[row + 2, 2].Style.Fill.PatternType = ExcelFillStyle.Solid;
                //worksheet.Cells[row+2, 2].Style.Fill.BackgroundColor.SetColor(Cores.DarkGreenWarning);
            }
            else if (input[1] == "LightGreen")
            {
                worksheet.Cells[row +2, 2].Value = "Medium";
                //worksheet.Cells[row + 2, 2].Style.Fill.PatternType = ExcelFillStyle.Solid;
                //worksheet.Cells[row + 2, 2].Style.Fill.BackgroundColor.SetColor(Cores.LightGreenWarning);
            }
            else if (input[1] == "Yellow")
            {
                worksheet.Cells[row + 2, 2].Value = "Low";
                //worksheet.Cells[row + 2, 2].Style.Fill.PatternType = ExcelFillStyle.Solid;
                //worksheet.Cells[row + 2, 2].Style.Fill.BackgroundColor.SetColor(Cores.YellowWarning);
            }
            else if (input[1] == "Orange")
            {
                worksheet.Cells[row + 2, 2].Value = "Medium";
                //worksheet.Cells[row + 2, 2].Style.Fill.PatternType = ExcelFillStyle.Solid;
                //worksheet.Cells[row + 2, 2].Style.Fill.BackgroundColor.SetColor(Cores.OrangeWarning);
            }
            else if (input[1] == "Red")
            {
                worksheet.Cells[row + 2, 2].Value = "High";
                //worksheet.Cells[row + 2, 2].Style.Fill.PatternType = ExcelFillStyle.Solid;
                //worksheet.Cells[row + 2, 2].Style.Fill.BackgroundColor.SetColor(Cores.RedWarning);
            }
        }
    }
}
