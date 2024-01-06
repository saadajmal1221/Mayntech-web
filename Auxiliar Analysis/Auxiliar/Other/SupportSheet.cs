using Mayntech___Individual_Solution.Auxiliar.Auxiliar.FMP;
using Mayntech___Individual_Solution.Auxiliar.Auxiliar.Other;
using Mayntech___Individual_Solution.Pages;
using OfficeOpenXml;
using OfficeOpenXml.Style;
using System.Drawing;

namespace Mayntech___Individual_Solution.Auxiliar_Analysis.Auxiliar.Other
{
    public class SupportSheet
    {

        public void supportConstr(ExcelPackage package, Dictionary<string, string> competitorsDict)
        {
            var workSheet = package.Workbook.Worksheets.Add("Support");
            workSheet.View.ShowGridLines = false;
            workSheet.View.ZoomScale = 80;
            workSheet.TabColor = Cores.corSecundária;

            // Cria a coluna azul em cima da tabela
            workSheet.Cells[2, 2].Value = "Competitors";
            workSheet.Cells[2, 2].Style.Font.Bold = true;
            workSheet.Cells[2, 2].Style.Font.Color.SetColor(Color.White);

            workSheet.Cells[3, 2].Value = "Symbol";
            workSheet.Cells[3, 2].Style.Font.Bold = true;
            workSheet.Cells[3, 2].Style.Font.Color.SetColor(Cores.CorTexto);
            workSheet.Cells[3, 2].Style.Border.Bottom.Style = ExcelBorderStyle.Thin;

            workSheet.Cells[3, 3].Value = "Company Name";
            workSheet.Cells[3, 3].Style.Font.Bold = true;
            workSheet.Cells[3, 3].Style.Font.Color.SetColor(Cores.CorTexto);
            workSheet.Cells[3, 3].Style.Border.Bottom.Style = ExcelBorderStyle.Thin;


            workSheet.Cells[2, 2].Style.Fill.PatternType = ExcelFillStyle.Solid;
            workSheet.Cells[2, 2].Style.Fill.BackgroundColor.SetColor(Cores.CorPrincipal);
            workSheet.Cells[2, 3].Style.Fill.PatternType = ExcelFillStyle.Solid;
            workSheet.Cells[2, 3].Style.Fill.BackgroundColor.SetColor(Cores.CorPrincipal);
            int aux = 1;

            foreach (KeyValuePair<string, string> item in competitorsDict)
            {
                try
                {
                    
                    workSheet.Cells[3 + aux, 2].Value = item.Value;
                    workSheet.Cells[3 + aux, 3].Value = item.Key;
                    aux++;
                }
                catch (Exception)
                {

                    continue;
                }


            }

            workSheet.Columns[3].Width = 35;

            workSheet.Cells[5 + aux, 2].Value = "Competitors Notes:";
            workSheet.Cells[5 + aux, 2].Style.Font.Bold = true;
            workSheet.Cells[6 + aux, 2].Value = "To assure comparability of the ratios, when needed, the financial statements of the competitors were adjusted to the date of the reference company.";
        }
    }
}
