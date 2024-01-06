using OfficeOpenXml.Style;
using OfficeOpenXml;
using System.Drawing;
using Mayntech___Individual_Solution.Auxiliar.Auxiliar.Other;

namespace Mayntech___Individual_Solution.Auxiliar.Analysis
{
    public class Disclaimer
    {
        public async Task CreateDisclaimer(ExcelPackage package)
        {
            var workSheet = package.Workbook.Worksheets.Add("Disclaimer");
            workSheet.View.ShowGridLines = false;
            workSheet.TabColor = Color.Gray;



            // Cria a coluna azul em cima da tabela
            workSheet.Cells[2, 2].Value = "Disclaimer";
            workSheet.Cells[2, 2].Style.Font.Bold = true;
            workSheet.Cells[2, 2].Style.Font.Color.SetColor(Color.White);


            workSheet.Cells["B2"].Style.Fill.PatternType = ExcelFillStyle.Solid;
            workSheet.Cells["B2"].Style.Fill.BackgroundColor.SetColor(Cores.CorPrincipal);

            //criar um field com o valor ser em milhoes ou thousands

            workSheet.Cells[3, 2].Value = "The goal of this report is to support the analysis of the company, but we highly recommend that the user performs additional research by looking at the company's annual reports and other relevant information. This report is for informational purposes only and should not be considered as investment advice. We are not responsible for the accuracy of the information provided and recommend users to verify it with other sources. Additionally, please note that the financial statements have been uniformized and there may be inaccurate allocations. Our sources are: Financial Modeling Prep (https://site.financialmodelingprep.com/) and Damodaran Online (https://pages.stern.nyu.edu/~adamodar/).";
            workSheet.Cells[3, 2].Style.WrapText = true;
            workSheet.Columns[2].Width = 80;
        }
    }
}
