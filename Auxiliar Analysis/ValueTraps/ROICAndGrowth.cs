using Mayntech___Individual_Solution.Auxiliar.Analysis;
using Mayntech___Individual_Solution.Auxiliar.Auxiliar.FMP;
using Mayntech___Individual_Solution.Auxiliar.Auxiliar.Other;
using Mayntech___Individual_Solution.Pages;
using OfficeOpenXml;
using OfficeOpenXml.Style;
using System.Drawing;

namespace Mayntech___Individual_Solution.Auxiliar_Analysis.ValueTraps
{
    public class ROICAndGrowth
    {
        public void roicContruction(ExcelPackage package, int numberOfYears, IDictionary< string,List<FinancialStatements>> incomeStatement,
            IDictionary<string, List<FinancialStatements>> balanceSheet, string companyName, string calendarYear)
        {
            ExcelNextCol columnName = new ExcelNextCol();
            var worksheet = package.Workbook.Worksheets.Add("ROIC analysis");
            worksheet.View.ShowGridLines = false;
            worksheet.View.ZoomScale = 80;
            worksheet.View.FreezePanes(4, 3);
            int row = 2;
            int col = 2;

            for (int i = 0; i < numberOfYears + 1; i++)
            {

                worksheet.Cells[row, col + i].Style.Fill.PatternType = ExcelFillStyle.Solid;
                worksheet.Cells[row, col + i].Style.Fill.BackgroundColor.SetColor(Cores.CorPrincipal);

                //worksheet.Cells[row + 20, col + i].Style.Fill.PatternType = ExcelFillStyle.Solid;
                //worksheet.Cells[row + 20, col + i].Style.Fill.BackgroundColor.SetColor(Cores.CorPrincipal);




                if (i == 0)
                {
                    worksheet.Cells[row, col].Value = "ROIC";
                    worksheet.Cells[row, col].Style.Font.Bold = true;
                    worksheet.Cells[row, col].Style.Font.Color.SetColor(Color.White);

                    worksheet.Cells[row + 1, col + i].Value = "Description";
                    worksheet.Cells[row + 1, col + i].Style.Font.Color.SetColor(Cores.CorTexto);
                    worksheet.Cells[row + 1, col + i].Style.Border.Bottom.Style = ExcelBorderStyle.Thin;
                    worksheet.Cells[row + 1, col + i].Style.Font.Bold = true;

                    worksheet.Cells[row + 3, col + i].Value = companyName;
                    worksheet.Cells[row + 4, col + i].Value = "Competitors";




                    worksheet.Columns[col + i].Width = 25;
                }
                else if (i < numberOfYears)
                {
                    string column = columnName.GetExcelColumnName(col + i);
                    string columnLeft = columnName.GetExcelColumnName(col + i - 1);
                    int year = int.Parse(calendarYear);
                    Dictionary<int, double> ROIC = GetPeersRoic(numberOfYears, incomeStatement, balanceSheet, year);
                    worksheet.Cells[row + 1, col + i].Formula = "=YEAR('P&L'!" + column + "3)";
                    worksheet.Cells[row + 1, col + i].Style.Font.Color.SetColor(Cores.CorTexto);
                    worksheet.Cells[row + 1, col + i].Style.Border.Bottom.Style = ExcelBorderStyle.Thin;
                    worksheet.Cells[row + 1, col + i].Style.Font.Bold = true;

                    worksheet.Rows[row + 2].Height = 3;

                    worksheet.Cells[row + 3, col + i].Formula = "=('Valuation Support'!" + column + "20/'Valuation Support'!" + column + "45)";
                    worksheet.Cells[row + 3, col + i].Style.Numberformat.Format = "0.00%";
                    try
                    {
                        worksheet.Cells[row + 4, col + i].Value = ROIC[year -numberOfYears+ i];
                        worksheet.Cells[row + 4, col + i].Style.Numberformat.Format = "0.00%";
                    }
                    catch 
                    {
                        continue;
                    }

                    


                }
            }
        }

        public Dictionary<int,double> GetPeersRoic(int numberOfYears, IDictionary<string, List<FinancialStatements>> incomeStatement,
            IDictionary<string, List<FinancialStatements>> balanceSheet, int CalendarYear)
        {
            List<double> ROICindividual = new List<double>();
            Dictionary<int, double> ROIC = new Dictionary<int, double>(); 
            for (int i = 0; i < Math.Min(5, numberOfYears); i++)
            {
                int year = 0;

                ROICindividual.Clear();

                foreach (var item in incomeStatement)
                {
                    try
                    {
                        if (item.Value[i].Date.Year == CalendarYear-i)
                        {

                            year = item.Value[i].Date.Year;

                            year = item.Value[i].Date.Year;
                            double noplat = item.Value[i].OperatingIncome;

                            //Estou a assumir tax rate de 20% aqui
                            noplat -= noplat * 0.2;

                            double NWC = balanceSheet[item.Key][i].NetReceivables + balanceSheet[item.Key][i].Inventory +
                                balanceSheet[item.Key][i].OtherCurrentAssets - balanceSheet[item.Key][i].AccountPayables -
                                balanceSheet[item.Key][i].DeferredRevenue - balanceSheet[item.Key][i].OtherCurrentLiabilities;

                            double investedCapital = balanceSheet[item.Key][i].propertyPlantEquipmentNet +
                                balanceSheet[item.Key][i].IntangibleAssets + balanceSheet[item.Key][i].Goodwill + NWC;

                            ROICindividual.Add(noplat / investedCapital);
                        }
                        else if (item.Value[i-1].Date.Year == CalendarYear-i)
                        {

                            year = item.Value[i-i].Date.Year;

                            year = item.Value[i-i].Date.Year;
                            double noplat = item.Value[i-i].OperatingIncome;

                            //Estou a assumir tax rate de 20% aqui
                            noplat -= noplat * 0.2;

                            double NWC = balanceSheet[item.Key][i-1].NetReceivables + balanceSheet[item.Key][i-1].Inventory +
                                balanceSheet[item.Key][i-1].OtherCurrentAssets - balanceSheet[item.Key][i-1].AccountPayables -
                                balanceSheet[item.Key][i-1].DeferredRevenue - balanceSheet[item.Key][i-1].OtherCurrentLiabilities;

                            double investedCapital = balanceSheet[item.Key][i-1].propertyPlantEquipmentNet +
                                balanceSheet[item.Key][i-1].IntangibleAssets + balanceSheet[item.Key][i-1].Goodwill + NWC;

                            ROICindividual.Add(noplat / investedCapital);
                        }
                        else if (item.Value[i + 1].Date.Year == CalendarYear-i)
                        {

                            year = item.Value[i + 1].Date.Year;

                            year = item.Value[i + 1].Date.Year;
                            double noplat = item.Value[i + 1].OperatingIncome;

                            double NWC = balanceSheet[item.Key][i + 1].NetReceivables + balanceSheet[item.Key][i + 1].Inventory +
                                balanceSheet[item.Key][i + 1].OtherCurrentAssets - balanceSheet[item.Key][i + 1].AccountPayables -
                                balanceSheet[item.Key][i + 1].DeferredRevenue - balanceSheet[item.Key][i + 1].OtherCurrentLiabilities;

                            double investedCapital = balanceSheet[item.Key][i + 1].propertyPlantEquipmentNet +
                                balanceSheet[item.Key][i + 1].IntangibleAssets + balanceSheet[item.Key][i + 1].Goodwill + NWC;

                            ROICindividual.Add(noplat / investedCapital);
                        
                        }
                        else
                        {

                        }

                    }
                    catch
                    {

                        continue;
                    }

                }


                if (ROICindividual.Count()>0)
                {
                    ROIC.Add(CalendarYear - i, ROICindividual.Average());
                }
                else
                {
                    ROIC.Add(CalendarYear - i, 0);
                }
            }
            return ROIC;
        }
    }

    public class IndustryDamodaran
    {
        public string CompanyName { get; set; }
        public string Exchange { get; set; }
        public string Ticker { get; set; }
        public string IndustryGroup { get; set; }

        public string PrimarySector { get; set; }
        public string CountryIso_2 { get; set; }
        public string Country { get; set; }
        public string BroadGroup { get; set; }
        public string SubGroup { get; set; }
    }
}
