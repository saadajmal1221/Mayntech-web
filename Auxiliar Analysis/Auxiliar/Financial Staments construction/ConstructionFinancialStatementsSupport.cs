using Mayntech___Individual_Solution.Auxiliar.Auxiliar.FMP;
using Mayntech___Individual_Solution.Auxiliar.Auxiliar.Other;
using OfficeOpenXml;
using OfficeOpenXml.Style;
using System.Drawing;


namespace Mayntech___Individual_Solution.Auxiliar.Auxiliar
{
    public class ConstructionFinancialStatementsSupport
    {

        public void CommonCaption(string description, double value, int col, int row, ExcelWorksheet workSheet, FinancialStatements item)
        {
            workSheet.Cells[row, 2].Value = description;
            workSheet.Cells[row, 3 + col].Value = value / 1000;
            workSheet.Cells[row, 3 + col].Style.Numberformat.Format = "#,##0;(#,##0);-";

        }
        public void CommonCaptionNegative(string description, double value, int col, int row, ExcelWorksheet workSheet, FinancialStatements item)
        {
            workSheet.Cells[row, 2].Value = description;
            workSheet.Cells[row, 3 + col].Value = value / 1000 * -1;
            workSheet.Cells[row, 3 + col].Style.Numberformat.Format = "#,##0;(#,##0);-";

        }
        public void CommonSubCaption(string description, double value, int col, int row, ExcelWorksheet workSheet, FinancialStatements item)
        {
            workSheet.Cells[row, 2].Value = description;
            workSheet.Cells[row, 2].Style.Fill.PatternType = ExcelFillStyle.Solid;
            workSheet.Cells[row, 2].Style.Fill.BackgroundColor.SetColor(Cores.corSecundária);
            workSheet.Cells[row, 3 + col].Value = value / 1000 * -1;
            workSheet.Cells[row, 3 + col].Style.Numberformat.Format = "#,##0;(#,##0);-";
            workSheet.Cells[row, 2].Style.Indent = 2;
            workSheet.Cells[row, 3 + col].Style.Fill.PatternType = ExcelFillStyle.Solid;
            workSheet.Cells[row, 3 + col].Style.Fill.BackgroundColor.SetColor(Cores.corSecundária);
            workSheet.Cells[row, 3 + col].Style.HorizontalAlignment = ExcelHorizontalAlignment.Right;

        }
        public void CaptionTotal(string description, double value, int col, int row, ExcelWorksheet workSheet, FinancialStatements item)
        {
            workSheet.Cells[row, 2].Value = description;
            workSheet.Cells[row, 3 + col].Value = value / 1000;
            workSheet.Cells[row, 3 + col].Style.Numberformat.Format = "#,##0;(#,##0);-";
            workSheet.Cells[row, 2].Style.Font.Bold = true;
            workSheet.Cells[row, 3 + col].Style.Font.Bold = true;

        }

        public void CaptionRatio(string description, double value, int col, int row, ExcelWorksheet workSheet, FinancialStatements item)
        {
            workSheet.Cells[row, 2].Value = description;
            workSheet.Cells[row, 3 + col].Value = value;

            workSheet.Cells[row, 3 + col].Style.Numberformat.Format = "0.00%";
            workSheet.Cells[row, 3 + col].Style.Font.Color.SetColor(Color.Gray);

        }

        //Corrigir a cor do texto para uma definida por nós (Cores.CorTexto)
        public void Divisor(string description, int col, int row, ExcelWorksheet workSheet)
        {
            workSheet.Cells[row, 2].Value = description;
            workSheet.Cells[row, 2].Style.Font.Color.SetColor(Cores.CorTexto2);
            workSheet.Cells[row, 2].Style.Font.Bold = true;

        }
        public void Subtitle(string description, int col, int row, ExcelWorksheet workSheet)
        {
            workSheet.Cells[row, 2].Value = description;
            workSheet.Cells[row, 2].Style.Font.Bold = true;
            workSheet.Cells[row, 2].Style.Indent = 1;


        }

    }
}
