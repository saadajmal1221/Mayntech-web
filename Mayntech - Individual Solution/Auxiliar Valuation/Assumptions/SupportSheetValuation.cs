using Mayntech___Individual_Solution.Auxiliar.Auxiliar.Other;
using OfficeOpenXml.Style;
using OfficeOpenXml;
using System.Drawing;
using Microsoft.EntityFrameworkCore.Query.Internal;

namespace Mayntech___Individual_Solution.Auxiliar_Valuation.Assumptions
{
    public class SupportSheetValuation
    {
        public void supportConstrVal(ExcelPackage package, Dictionary<string, string> competitorsDict)
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


            // Cria a coluna azul em cima da tabela
            workSheet.Cells[2, 5].Value = "Sheet Assumptions";
            workSheet.Cells[2, 5].Style.Font.Bold = true;
            workSheet.Cells[2, 5].Style.Font.Color.SetColor(Color.White);

            workSheet.Cells[2, 5].Style.Fill.PatternType = ExcelFillStyle.Solid;
            workSheet.Cells[2, 5].Style.Fill.BackgroundColor.SetColor(Cores.CorPrincipal);

            workSheet.Cells[3, 5].Style.Font.Bold = true;
            workSheet.Cells[3, 5].Style.Font.Color.SetColor(Cores.CorTexto);
            workSheet.Cells[3, 5].Style.Border.Bottom.Style = ExcelBorderStyle.Thin;

            string auxAssumption = '"' + "Assumptions" + '"';
            string auxOtherAssumptions = '"' + "Other Assumptions" + '"';

            workSheet.Cells[4, 5].Value = "The sheet " + auxAssumption + "contains all the assumptions used for the DCF calculations.";
            workSheet.Cells[5, 5].Value = "While all the assumptions in this page should be questioned, in the cells C5:E10, there are inputs that are supposed to be filled by the user, as illustrated in cell C5. ";
            workSheet.Cells[6, 5].Value = "This inputs should reflect the user's forecast for the given captions. ";
            workSheet.Cells[7, 5].Value = "We recommend that this input's are support by an extensive examination of the company, industry and the economy as a whole.";

            workSheet.Cells[9, 5].Value = "The inputs present in " +auxOtherAssumptions + " are a combination of Financial Modeling Prep and Damodaran assumptions.";

            workSheet.Columns[5].Width = 167;

        }
    }
}

