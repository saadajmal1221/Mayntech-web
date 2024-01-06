using Mayntech___Individual_Solution.Auxiliar.Auxiliar.FMP;
using Mayntech___Individual_Solution.Auxiliar.Auxiliar.Other;
using Mayntech___Individual_Solution.Pages;
using OfficeOpenXml;
using OfficeOpenXml.Style;
using System.Drawing;

namespace Mayntech___Individual_Solution.Auxiliar.Analysis
{
    public class CompanyOverview
    {
        public void CompanyOverviewConstruction(CompanyProfile companyOutlook, ExcelPackage package)
        {
            var workSheet = package.Workbook.Worksheets.Add("Company Overview");
            workSheet.View.ShowGridLines = false;
            workSheet.View.ZoomScale = 100;


            workSheet.Cells[2, 2].Value = "Company Overview";
            workSheet.Cells[2, 2].Style.Font.Bold = true;
            workSheet.Cells[2, 2].Style.Font.Color.SetColor(Color.White);

            workSheet.Cells["B2"].Style.Fill.PatternType = ExcelFillStyle.Solid;
            workSheet.Cells["B2"].Style.Fill.BackgroundColor.SetColor(Cores.CorPrincipal);

            workSheet.Cells["C2"].Style.Fill.PatternType = ExcelFillStyle.Solid;
            workSheet.Cells["C2"].Style.Fill.BackgroundColor.SetColor(Cores.CorPrincipal);

            int row = 3;
            workSheet.Cells[row, 2].Value = "Company Name";
            workSheet.Cells[row, 3].Value = companyOutlook.profile.companyName;
            workSheet.Cells[row, 3].Style.Font.Bold = true;
            row++;

            workSheet.Cells[row, 2].Value = "Company ticker";
            workSheet.Cells[row, 3].Value = companyOutlook.profile.symbol;
            workSheet.Cells[row, 3].Style.Font.Bold = true;
            row++;

            workSheet.Cells[row, 2].Value = "ISIN";
            workSheet.Cells[row, 3].Value = companyOutlook.profile.isin;
            workSheet.Cells[row, 3].Style.Font.Bold = true;
            row++;

            workSheet.Cells[row, 2].Value = "Country";
            workSheet.Cells[row, 3].Value = companyOutlook.profile.country;
            workSheet.Cells[row, 3].Style.Font.Bold = true;
            row++;

            workSheet.Cells[row, 2].Value = "Currency";
            workSheet.Cells[row, 3].Value = companyOutlook.profile.currency;
            workSheet.Cells[row, 3].Style.Font.Bold = true;
            row++;

            workSheet.Cells[row, 2].Value = "Exchange";
            workSheet.Cells[row, 3].Value = companyOutlook.profile.exchangeShortName;
            workSheet.Cells[row, 3].Style.Font.Bold = true;
            row++;

            workSheet.Cells[row, 2].Value = "Sector";
            workSheet.Cells[row, 3].Value = companyOutlook.profile.sector;
            workSheet.Cells[row, 3].Style.Font.Bold = true;
            row++;

            workSheet.Cells[row, 2].Value = "Industry";
            workSheet.Cells[row, 3].Value = companyOutlook.profile.industry;
            workSheet.Cells[row, 3].Style.Font.Bold = true;
            row++;

            workSheet.Cells[row, 2].Value = "CEO";
            workSheet.Cells[row, 3].Value = companyOutlook.profile.ceo;
            workSheet.Cells[row, 3].Style.Font.Bold = true;
            row++;

            workSheet.Cells[row, 2].Value = "Full time employees";
            workSheet.Cells[row, 3].Value = companyOutlook.profile.fullTimeEmployees;
            workSheet.Cells[row, 3].Style.Numberformat.Format = "#,##0;(#,##0);-";
            workSheet.Cells[row, 3].Style.Font.Bold = true;
            row++;

            workSheet.Cells[row, 2].Value = "Market Capitalization";
            workSheet.Cells[row, 3].Value = companyOutlook.profile.mktCap;
            workSheet.Cells[row, 3].Style.Numberformat.Format = "#,##0;(#,##0);-";
            workSheet.Cells[row, 3].Style.Font.Bold = true;
            row++;

            workSheet.Cells[row, 2].Value = "Share Price";
            workSheet.Cells[row, 3].Value = companyOutlook.profile.price;
            workSheet.Cells[row, 3].Style.Font.Bold = true;
            row++;

            workSheet.Columns[3].AutoFit();
            workSheet.Columns[2].AutoFit();

            workSheet.Cells["E2"].Value = "Description";
            workSheet.Cells["E2"].Style.Font.Bold = true;
            workSheet.Cells["E2"].Style.Font.Color.SetColor(Color.White);

            workSheet.Cells["E2:K2"].Style.Fill.PatternType = ExcelFillStyle.Solid;
            workSheet.Cells["E2:K2"].Style.Fill.BackgroundColor.SetColor(Cores.CorPrincipal);

            workSheet.Cells[3,5].Value = companyOutlook.profile.description;

            workSheet.Cells["E3:K14"].Merge = true;
            workSheet.Cells["E3:K14"].Style.VerticalAlignment = ExcelVerticalAlignment.Top; ;
            workSheet.Cells["E3:K14"].Style.WrapText = true;

            try
            {
                workSheet.Cells[row + 1, 2].Value = "Report";
                var cell = workSheet.Cells[row + 1, 3];
                cell.Hyperlink = new Uri(SolutionModel.incomeStatement[0].FinalLink);
                workSheet.Cells[row + 1, 3].Value = SolutionModel.incomeStatement[0].FinalLink;
            }
            catch 
            {


            }




            row++;

        }
    }
}
