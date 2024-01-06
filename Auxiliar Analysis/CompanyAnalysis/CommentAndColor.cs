using Mayntech___Individual_Solution.Auxiliar.Auxiliar.Other;
using OfficeOpenXml;
using OfficeOpenXml.Style;
using System.Drawing;

namespace Mayntech___Individual_Solution.Auxiliar.CompanyAnalysis
{
    public class CommentAndColor : CalculationsAux
    {
        public void Comment(ExcelWorksheet worksheet, int row, int column, List<int> color, List<string> PositiveComment, List<string> NegativeComment, List<string> OtherComment)
        {
            if (color.Count()>0)
            {
                int FinalColor = color.Max();
                int MinColor = color.Min();
                if (FinalColor == 5)
                {
                    worksheet.Cells[row, column].Style.Fill.PatternType = ExcelFillStyle.Solid;
                    worksheet.Cells[row, column].Style.Fill.BackgroundColor.SetColor(Cores.RedWarning);
                }
                else if (FinalColor == 4)
                {
                    if (color.Average()<3)
                    {
                        worksheet.Cells[row, column].Style.Fill.PatternType = ExcelFillStyle.Solid;
                        worksheet.Cells[row, column].Style.Fill.BackgroundColor.SetColor(Cores.OrangeWarning);
                    }
                    else
                    {
                        worksheet.Cells[row, column].Style.Fill.PatternType = ExcelFillStyle.Solid;
                        worksheet.Cells[row, column].Style.Fill.BackgroundColor.SetColor(Cores.YellowWarning);
                    }

                }
                else if (FinalColor == 3)
                {
                    if (MinColor == 1)
                    {
                        worksheet.Cells[row, column].Style.Fill.PatternType = ExcelFillStyle.Solid;
                        worksheet.Cells[row, column].Style.Fill.BackgroundColor.SetColor(Cores.LightGreenWarning);
                    }
                    else if (color.Average() <2.5)
                    {
                        worksheet.Cells[row, column].Style.Fill.PatternType = ExcelFillStyle.Solid;
                        worksheet.Cells[row, column].Style.Fill.BackgroundColor.SetColor(Cores.LightGreenWarning);
                    }
                    worksheet.Cells[row, column].Style.Fill.PatternType = ExcelFillStyle.Solid;
                    worksheet.Cells[row, column].Style.Fill.BackgroundColor.SetColor(Cores.YellowWarning);
                }
                else if (FinalColor == 2)
                {
                    if (MinColor == 1)
                    {
                        worksheet.Cells[row, column].Style.Fill.PatternType = ExcelFillStyle.Solid;
                        worksheet.Cells[row, column].Style.Fill.BackgroundColor.SetColor(Cores.DarkGreenWarning);
                    }
                    worksheet.Cells[row, column].Style.Fill.PatternType = ExcelFillStyle.Solid;
                    worksheet.Cells[row, column].Style.Fill.BackgroundColor.SetColor(Cores.LightGreenWarning);
                }
                else if (FinalColor == 1)
                {
                    worksheet.Cells[row, column].Style.Fill.PatternType = ExcelFillStyle.Solid;
                    worksheet.Cells[row, column].Style.Fill.BackgroundColor.SetColor(Cores.DarkGreenWarning);
                }


            }
            else
            {
                worksheet.Cells[row, column].Style.Fill.PatternType = ExcelFillStyle.Solid;
                worksheet.Cells[row, column].Style.Fill.BackgroundColor.SetColor(Color.Gray);
            }

            var boldRichText = worksheet.Cells[row, column + 1].RichText.Add("Positive: ");
            boldRichText.Bold = true;

            if (PositiveComment.Count() != 0)
            {
                var stringPositive = string.Join("", PositiveComment);
                var normalRichText = worksheet.Cells[row, column + 1].RichText.Add(stringPositive);
                normalRichText.Bold = false;
            }
            else
            {
                var stringPositive = "N/A. ";
                var normalRichText = worksheet.Cells[row, column + 1].RichText.Add(stringPositive);
                normalRichText.Bold = false;
            }


            var boldRichNegative = worksheet.Cells[row, column + 1].RichText.Add("Negative: ");
            boldRichNegative.Bold = true;

            if (NegativeComment.Count() != 0)
            {
                var Negative = string.Join("", NegativeComment);
                var normalRichNegative = worksheet.Cells[row, column + 1].RichText.Add(Negative);
                normalRichNegative.Bold = false;
            }
            else
            {
                var Negative = "N/A. ";
                var normalRichText = worksheet.Cells[row, column + 1].RichText.Add(Negative);
                normalRichText.Bold = false;
            }




            if (OtherComment.Count()>0)
            {
                var boldRichOther = worksheet.Cells[row, column + 1].RichText.Add("Other: ");
                boldRichOther.Bold = true;

                var Other = string.Join("", OtherComment);
                var normalRichOther = worksheet.Cells[row, column + 1].RichText.Add(Other);
                normalRichOther.Bold = false;
            }

            worksheet.Cells[row, column + 1].Style.VerticalAlignment = ExcelVerticalAlignment.Top; ;
            worksheet.Cells[row, column + 1].Style.WrapText = true;
            
        }
    }

}
