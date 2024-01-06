using Mayntech___Individual_Solution.Auxiliar.Auxiliar.Other;
using Mayntech___Individual_Solution.Pages;
using OfficeOpenXml;
using OfficeOpenXml.Style;
using System.Drawing;

namespace Mayntech___Individual_Solution.Auxiliar_Analysis.ValueTraps
{
    public class AuxiliarAnalysis
    {
        public void AuxiliarConstruction(ExcelPackage package, int numberOfYears)
        {
            ExcelNextCol columnName = new ExcelNextCol();
            var worksheet = package.Workbook.Worksheets.Add("Auxiliar");
            worksheet.View.ShowGridLines = false;
            worksheet.View.ZoomScale = 80;
            worksheet.View.FreezePanes(4, 3);
            int row = 2;
            int col = 2;


            for (int i = 0; i < numberOfYears +1; i++)
            {

                worksheet.Cells[row, col + i].Style.Fill.PatternType = ExcelFillStyle.Solid;
                worksheet.Cells[row, col + i].Style.Fill.BackgroundColor.SetColor(Cores.CorPrincipal);



                //worksheet.Cells[row + 20, col + i].Style.Fill.PatternType = ExcelFillStyle.Solid;
                //worksheet.Cells[row + 20, col + i].Style.Fill.BackgroundColor.SetColor(Cores.CorPrincipal);




                if (i == 0)
                {
                    worksheet.Cells[row, col].Value = "Cash Conversion cycle";
                    worksheet.Cells[row, col].Style.Font.Bold = true;
                    worksheet.Cells[row, col].Style.Font.Color.SetColor(Color.White);

                    worksheet.Cells[row + 1, col + i].Value = "Description";
                    worksheet.Cells[row + 1, col + i].Style.Font.Color.SetColor(Cores.CorTexto);
                    worksheet.Cells[row + 1, col + i].Style.Border.Bottom.Style = ExcelBorderStyle.Thin;
                    worksheet.Cells[row + 1, col + i].Style.Font.Bold = true;

                    worksheet.Cells[row + 3, col + i].Value = "DPO";
                    worksheet.Cells[row + 4, col + i].Value = "DSO";
                    worksheet.Cells[row + 5, col + i].Value = "DIO";
                    worksheet.Cells[row + 6, col + i].Value = "CCC";






                    worksheet.Cells[row + 10, col + i].Value = "Depreciation & Amortization";
                    worksheet.Cells[row + 10, col + i].Style.Font.Color.SetColor(Cores.CorTexto);
                    worksheet.Cells[row + 10, col + i].Style.Border.Bottom.Style = ExcelBorderStyle.Thin;
                    worksheet.Cells[row + 10, col + i].Style.Font.Bold = true;
                    worksheet.Cells[row + 11, col + i].Value = "DP&A";
                    worksheet.Cells[row + 12, col + i].Value = "DP&A rate";
                    worksheet.Cells[row + 12, col + i].Style.Font.Color.SetColor(Cores.corSecundáriaText);

                    

                    worksheet.Columns[col + i].Width = 25;
                }
                else if (i < numberOfYears + 1)
                {
                    string column = columnName.GetExcelColumnName(col + i);
                    string columnLeft = columnName.GetExcelColumnName(col + i - 1);
                    worksheet.Cells[row + 1, col + i].Formula = "=YEAR('P&L'!" + column + "3)";
                    worksheet.Cells[row + 1, col + i].Style.Font.Color.SetColor(Cores.CorTexto);
                    worksheet.Cells[row + 1, col + i].Style.Border.Bottom.Style = ExcelBorderStyle.Thin;
                    worksheet.Cells[row + 1, col + i].Style.Font.Bold = true;

                    worksheet.Rows[row + 2].Height = 3;

                    worksheet.Cells[row + 3, col + i].Formula = "=('BS'!" + column + "27/-'P&L'!" + column + "6)*365";
                    worksheet.Cells[row + 3, col + i].Style.Numberformat.Format = "0";
                    worksheet.Cells[row + 4, col + i].Formula = "=('BS'!" + column + "9/'P&L'!" + column + "5)*365";
                    worksheet.Cells[row + 4, col + i].Style.Numberformat.Format = "0";
                    worksheet.Cells[row + 5, col + i].Formula = "=('BS'!" + column + "10/-'P&L'!" + column + "6)*365";
                    worksheet.Cells[row + 5, col + i].Style.Numberformat.Format = "0";
                    worksheet.Cells[row + 6, col + i].Formula = "=-" + column + "5+" + column + "6+" + column + "7";
                    worksheet.Cells[row + 6, col + i].Style.Numberformat.Format = "0";



                    //DP&A
                    worksheet.Cells[row + 10, col + i].Style.Border.Bottom.Style = ExcelBorderStyle.Thin;
                    worksheet.Cells[row + 10, col + i].Style.Font.Bold = true;
                    worksheet.Cells[row + 11, col + i].Formula = "=-'P&L'!" + column + "32";
                    worksheet.Cells[row + 11, col + i].Style.Numberformat.Format = "#,##0;(#,##0);-";
                    worksheet.Cells[row + 12, col + i].Formula = "=-'P&L'!" + column + "32/('BS'!" + column + "15+'BS'!" + column + "17)";
                    worksheet.Cells[row + 12, col + i].Style.Numberformat.Format = "0.00%";
                    worksheet.Cells[row + 12, col + i].Style.Font.Color.SetColor(Cores.corSecundáriaText);


                    worksheet.Columns[col + i].AutoFit();

                }
                
            }
        }
    }
}
