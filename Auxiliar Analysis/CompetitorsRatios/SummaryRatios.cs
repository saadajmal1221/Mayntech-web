using Mayntech___Individual_Solution.Auxiliar.Auxiliar.Other;
using OfficeOpenXml;
using OfficeOpenXml.Style;
using System.Drawing;

namespace Mayntech___Individual_Solution.Auxiliar_Analysis.CompetitorsRatios
{
    public class SummaryRatios
    {
        public void summary(ExcelWorksheet worksheet, List<double> points, List<string> comments)
        {
            //Faz o exec.Summary da liquidity
            worksheet.Cells["B3:B4"].Merge = true;
            worksheet.Cells["B3:B4"].Style.VerticalAlignment = ExcelVerticalAlignment.Center;
            worksheet.Cells["B3:B4"].Style.HorizontalAlignment = ExcelHorizontalAlignment.Center;
            worksheet.Cells["B3:B4"].Style.WrapText = true;
            worksheet.Cells["B3"].Value = "Score";

            worksheet.Cells["C3:c4"].Merge = true;
            worksheet.Cells["C3:C4"].Style.VerticalAlignment = ExcelVerticalAlignment.Center;
            worksheet.Cells["C3:C4"].Style.HorizontalAlignment = ExcelHorizontalAlignment.Center;
            worksheet.Cells["C3:C4"].Style.WrapText = true;

            var modelTable = worksheet.Cells["B3:C4"];
            double sum = points.Sum();
            double score = sum / (points.Count());
            worksheet.Cells["c3"].Value = score;
            worksheet.Cells["c3"].Style.Numberformat.Format = "0%";

            if (score > 0.90)
            {
                worksheet.Cells["c3"].Style.Fill.PatternType = ExcelFillStyle.Solid;
                worksheet.Cells["c3"].Style.Fill.BackgroundColor.SetColor(Cores.DarkGreenWarning);
                worksheet.Cells["E3"].Value = comments[0];
            }
            else if (score > 0.50)
            {
                worksheet.Cells["c3"].Style.Fill.PatternType = ExcelFillStyle.Solid;
                worksheet.Cells["c3"].Style.Fill.BackgroundColor.SetColor(Cores.LightGreenWarning);
                worksheet.Cells["E3"].Value = comments[1];

            }
            else if (score > 0.10)
            {
                worksheet.Cells["c3"].Style.Fill.PatternType = ExcelFillStyle.Solid;
                worksheet.Cells["c3"].Style.Fill.BackgroundColor.SetColor(Cores.YellowWarning);
                worksheet.Cells["E3"].Value = comments[1];

            }
            else
            {
                worksheet.Cells["c3"].Style.Fill.PatternType = ExcelFillStyle.Solid;
                worksheet.Cells["c3"].Style.Fill.BackgroundColor.SetColor(Cores.OrangeWarning);
                worksheet.Cells["E3"].Value = comments[2];
            }
            points.Clear();




            modelTable.Style.Border.Top.Style = ExcelBorderStyle.Thin;
            modelTable.Style.Border.Left.Style = ExcelBorderStyle.Thin;
            modelTable.Style.Border.Right.Style = ExcelBorderStyle.Thin;
            modelTable.Style.Border.Bottom.Style = ExcelBorderStyle.Thin;

            worksheet.Row(7).Style.Fill.PatternType = ExcelFillStyle.Solid;
            worksheet.Row(7).Style.Fill.BackgroundColor.SetColor(Color.Gray);
            worksheet.Cells[7, 2].Value = "Ratios";
            worksheet.Cells[7, 2].Style.Font.Color.SetColor(Color.White);
            worksheet.Cells[7, 2].Style.Font.Bold = true;
        }
    }
}
