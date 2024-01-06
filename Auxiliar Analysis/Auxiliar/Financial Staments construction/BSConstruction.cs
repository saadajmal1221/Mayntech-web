using OfficeOpenXml.Style;
using OfficeOpenXml;
using System.Drawing;
using System.Globalization;
using System.Runtime.Versioning;
using Mayntech___Individual_Solution.Auxiliar.Auxiliar.FMP;
using Mayntech___Individual_Solution.Auxiliar.Auxiliar.Other;
using Microsoft.EntityFrameworkCore.Metadata.Internal;

namespace Mayntech___Individual_Solution.Auxiliar.Auxiliar
{
    public class BSConstruction
    {
        public int totAssets;
        public int totLiabilities;
        public int totEquity;
        public async Task CreateBS(ExcelPackage package, List<FinancialStatements> balances, int col, int row, string companyName,
            PeersProfile companyProfile)
        {
            var workSheet = package.Workbook.Worksheets.Add("BS");
            workSheet.View.ShowGridLines = false;
            workSheet.View.FreezePanes(4, 3);

            IncomeStatementConstruction incomeStatementConstruction = new IncomeStatementConstruction();

            int quarters = incomeStatementConstruction.Quarters(companyProfile);
            //Cria uma instance da class cores para fazer as cores de fundo e do texto
            //Cores cor = new Cores();
            //cor.CorPrincipal = System.Drawing.ColorTranslator.FromHtml("#576EDF");
            //cor.corSecundária = System.Drawing.ColorTranslator.FromHtml("#D9E1F2");
            //cor.CorNumber3 = System.Drawing.ColorTranslator.FromHtml("#C6E0B4");
            //cor.CorTexto = System.Drawing.ColorTranslator.FromHtml("#203764");

            // Cria a coluna azul em cima da tabela
            workSheet.Cells[2, 2].Value = "Balance Sheet - " + companyName;
            workSheet.Cells[2, 2].Style.Font.Bold = true;
            workSheet.Cells[2, 2].Style.Font.Color.SetColor(Color.White);

            // atenção que este "H2" está hardcoded. Tem de ser refeito para variar com o número de anos
            workSheet.Cells["B2"].Style.Fill.PatternType = ExcelFillStyle.Solid;
            workSheet.Cells["B2"].Style.Fill.BackgroundColor.SetColor(Cores.CorPrincipal);

            //criar um field com o valor ser em milhoes ou thousands
            ExcelNextCol columnName = new ExcelNextCol();

            balances.Reverse();
            foreach (FinancialStatements item in balances)
            {

                int row1 = 5;
                ConstructionFinancialStatementsSupport support = new();

                workSheet.Cells[2, 3 + col].Style.Fill.PatternType = ExcelFillStyle.Solid;
                workSheet.Cells[2, 3 + col].Style.Fill.BackgroundColor.SetColor(Cores.CorPrincipal);

                // cria o header
                workSheet.Cells[3, 2].Value = "Description (in '000 " + item.ReportedCurrency + ")";
                workSheet.Cells[3, 2].Style.Font.Bold = true;
                workSheet.Cells[3, 2].Style.Font.Color.SetColor(Cores.CorTexto);
                workSheet.Cells[3, 2].Style.Border.Bottom.Style = ExcelBorderStyle.Thin;

                // inserir a data com as formatações
                workSheet.Cells[row1 - 2, 3 + col].Value = item.Date;
                workSheet.Cells[row1 - 2, 3 + col].Style.Numberformat.Format = DateTimeFormatInfo.CurrentInfo.ShortDatePattern;
                workSheet.Cells[row1 - 2, 3 + col].Style.Font.Bold = true;
                workSheet.Cells[row1 - 2, 3 + col].Style.Font.Color.SetColor(Cores.CorTexto);
                workSheet.Cells[row1 - 2, 3 + col].Style.Border.Bottom.Style = ExcelBorderStyle.Thin;

                // define o height da linha entre o cabeçalho e os valores
                workSheet.Row(row1 - 1).Height = 4;

                string colAux = columnName.GetExcelColumnName(col + 3);


                support.Divisor("Assets", col, row1, workSheet);
                row1++;
                support.Subtitle("Current assets", col, row1, workSheet);
                row1++;

                if (item.CashAndCashEquivalents != null)
                {
                    support.CommonCaption("Cash and cash equivalents", (double)item.CashAndCashEquivalents, col, row1, workSheet, item);
                    workSheet.Cells[row1, 2].Style.Indent = 2;
                    row1++;
                }
                if (item.ShortTermInvestments != null)
                {
                    support.CommonCaption("Short term investments", (double)item.ShortTermInvestments, col, row1, workSheet, item);
                    workSheet.Cells[row1, 2].Style.Indent = 2;
                    row1++;
                }

                if (item.NetReceivables != null)
                {
                    support.CommonCaption("Accounts receivable, net", (double)item.NetReceivables, col, row1, workSheet, item);
                    workSheet.Cells[row1, 2].Style.Indent = 2;
                    row1++;
                }
                if (item.Inventory != null)
                {
                    support.CommonCaption("Inventory", (double)item.Inventory, col, row1, workSheet, item);
                    workSheet.Cells[row1, 2].Style.Indent = 2;
                    row1++;
                }
                if (item.OtherCurrentAssets != null)
                {
                    support.CommonCaption("Other current assets", (double)item.OtherCurrentAssets, col, row1, workSheet, item);
                    workSheet.Cells[row1, 2].Style.Indent = 2;
                    row1++;
                }
                if (item.TotalCurrentAssets != null)
                {
                    workSheet.Cells[row1, 2].Value = "Total current assets";
                    workSheet.Cells[row1, 2].Style.Font.Bold = true;
                    workSheet.Cells[row1, col + 3].Formula = "=" + colAux + "7+" + colAux + "8+" + colAux + "9+" + colAux + "10+" + colAux + "11";
                    workSheet.Cells[row1, col + 3].Style.Numberformat.Format = "#,##0;(#,##0);-";
                    workSheet.Cells[row1, col + 3].Style.Font.Bold = true;
                    
                    workSheet.Cells[row1, 2].Style.Indent = 2;
                    row1++;
                }
                row1++;
                support.Subtitle("Non-current assets", col, row1, workSheet);
                row1++;
                if (item.propertyPlantEquipmentNet != null)
                {
                    support.CommonCaption("Property plant and equipment", (double)item.propertyPlantEquipmentNet, col, row1, workSheet, item);
                    workSheet.Cells[row1, 2].Style.Indent = 2;
                    row1++;
                }
                if (item.Goodwill != null)
                {
                    support.CommonCaption("Goodwill", (double)item.Goodwill, col, row1, workSheet, item);
                    workSheet.Cells[row1, 2].Style.Indent = 2;
                    row1++;
                }
                if (item.IntangibleAssets != null)
                {
                    support.CommonCaption("Intangible assets", (double)item.IntangibleAssets, col, row1, workSheet, item);
                    workSheet.Cells[row1, 2].Style.Indent = 2;
                    row1++;
                }
                if (item.LongtermInvestments != null)
                {
                    support.CommonCaption("Long term investments", (double)item.LongtermInvestments, col, row1, workSheet, item);
                    workSheet.Cells[row1, 2].Style.Indent = 2;
                    row1++;
                }
                if (item.TaxAssets != null)
                {
                    support.CommonCaption("Tax assets", (double)item.TaxAssets, col, row1, workSheet, item);
                    workSheet.Cells[row1, 2].Style.Indent = 2;
                    row1++;
                }
                if (item.OtherNonCurrentAssets != null)
                {
                    support.CommonCaption("Other non-current assets", (double)item.OtherNonCurrentAssets, col, row1, workSheet, item);
                    workSheet.Cells[row1, 2].Style.Indent = 2;
                    row1++;
                }
                if (item.TotalNonCurrentassets != null)
                {
                    workSheet.Cells[row1, 2].Value = "Total non-current assets";
                    workSheet.Cells[row1, 2].Style.Font.Bold = true;
                    workSheet.Cells[row1, col + 3].Formula = "=" + colAux + "15+" + colAux + "16+" + colAux + "17+" + colAux + "18+" + colAux + "19+" + colAux + "20";
                    workSheet.Cells[row1, col + 3].Style.Numberformat.Format = "#,##0;(#,##0);-";
                    workSheet.Cells[row1, col + 3].Style.Font.Bold = true;
                    
                    workSheet.Cells[row1, 2].Style.Indent = 2;
                    row1++;
                    row1++;
                }
                if (item.TotalAssets != null)
                {
                    workSheet.Cells[row1, 2].Value = "Total assets";
                    workSheet.Cells[row1, 2].Style.Font.Bold = true;
                    workSheet.Cells[row1, col + 3].Formula = "=" + colAux + "12+" + colAux + "21";
                    workSheet.Cells[row1, col + 3].Style.Numberformat.Format = "#,##0;(#,##0);-";
                    workSheet.Cells[row1, col + 3].Style.Font.Bold = true;

                    workSheet.Cells[row1, 2].Style.Font.Color.SetColor(Cores.CorTexto2);
                    workSheet.Cells[row1, 3 + col].Style.Border.Bottom.Style = ExcelBorderStyle.Thin;
                    workSheet.Cells[row1, 2].Style.Border.Bottom.Style = ExcelBorderStyle.Thin;
                    totAssets = row1;
                    row1++;


                }
                row1++;
                support.Divisor("Liabilities", col, row1, workSheet);
                row1++;
                support.Subtitle("Current Liabilities", col, row1, workSheet);
                row1++;
                if (item.AccountPayables != null)
                {
                    support.CommonCaption("Accounts payable", (double)item.AccountPayables, col, row1, workSheet, item);
                    workSheet.Cells[row1, 2].Style.Indent = 2;
                    row1++;
                }
                if (item.ShortTermDebt != null)
                {
                    support.CommonCaption("Short term debt", (double)item.ShortTermDebt, col, row1, workSheet, item);
                    workSheet.Cells[row1, 2].Style.Indent = 2;
                    row1++;
                }
                if (item.TaxPayables != null)
                {
                    support.CommonCaption("Tax payables", 0, col, row1, workSheet, item);
                    workSheet.Cells[row1, 2].Style.Indent = 2;
                    row1++;
                }
                if (item.DeferredRevenue != null)
                {
                    support.CommonCaption("Deferred revenue", (double)item.DeferredRevenue, col, row1, workSheet, item);
                    workSheet.Cells[row1, 2].Style.Indent = 2;
                    row1++;
                }
                if (item.OtherCurrentLiabilities != null)
                {
                    support.CommonCaption("Other current liabilities", item.OtherCurrentLiabilities, col, row1, workSheet, item);
                    workSheet.Cells[row1, 2].Style.Indent = 2;
                    row1++;
                }
                if (item.totalCurrentLiabilities != null)
                {
                    workSheet.Cells[row1, 2].Value = "Total current liabilities";
                    workSheet.Cells[row1, 2].Style.Font.Bold = true;
                    workSheet.Cells[row1, col + 3].Formula = "=" + colAux + "27+" + colAux + "28+" + colAux + "29+" + colAux + "30+" + colAux + "31";
                    workSheet.Cells[row1, col + 3].Style.Numberformat.Format = "#,##0;(#,##0);-";
                    workSheet.Cells[row1, col + 3].Style.Font.Bold = true;
                    
                    workSheet.Cells[row1, 2].Style.Indent = 2;
                    row1++;
                    row1++;
                }
                support.Subtitle("Non-current Liabilities", col, row1, workSheet);
                row1++;
                if (item.LongTermDebt != null)
                {
                    support.CommonCaption("Long term debt", item.LongTermDebt, col, row1, workSheet, item);
                    workSheet.Cells[row1, 2].Style.Indent = 2;
                    row1++;
                }
                if (item.DeferredRevenueNonCurrent != null)
                {
                    support.CommonCaption("Deferred revenue, non-current", item.DeferredRevenueNonCurrent, col, row1, workSheet, item);
                    workSheet.Cells[row1, 2].Style.Indent = 2;
                    row1++;
                }
                if (item.DeferredTaxLiabilitiesNonCurrent != null)
                {
                    support.CommonCaption("Deferred tax liabilities, non-current", item.DeferredTaxLiabilitiesNonCurrent, col, row1, workSheet, item);
                    workSheet.Cells[row1, 2].Style.Indent = 2;
                    row1++;
                }
                if (item.OtherLiabilities != null)
                {
                    support.CommonCaption("Other liabilities", item.OtherLiabilities, col, row1, workSheet, item);
                    workSheet.Cells[row1, 2].Style.Indent = 2;
                    row1++;
                }
                if (item.OtherNonCurrentLiabilities != null)
                {
                    support.CommonCaption("Other non-current liabilities", item.OtherNonCurrentLiabilities, col, row1, workSheet, item);
                    workSheet.Cells[row1, 2].Style.Indent = 2;
                    row1++;
                }
                if (item.TotalNonCurrentLiabilities != null)
                {
                    workSheet.Cells[row1, 2].Value = "Total non-current liabilities";
                    workSheet.Cells[row1, 2].Style.Font.Bold = true;
                    workSheet.Cells[row1, col + 3].Formula = "=" + colAux + "35+" + colAux + "36+" + colAux + "37+" + colAux + "38+" + colAux + "39";
                    workSheet.Cells[row1, col + 3].Style.Numberformat.Format = "#,##0;(#,##0);-";
                    workSheet.Cells[row1, col + 3].Style.Font.Bold = true;
                    workSheet.Cells[row1, 2].Style.Indent = 2;
                    row1++;
                    row1++;
                }
                if (item.totalLiabilities != null)
                {
                    workSheet.Cells[row1, 2].Value = "Total liabilities";
                    workSheet.Cells[row1, 2].Style.Font.Bold = true;
                    workSheet.Cells[row1, col + 3].Formula = "=" + colAux + "32+" + colAux + "40";
                    workSheet.Cells[row1, col + 3].Style.Numberformat.Format = "#,##0;(#,##0);-";
                    workSheet.Cells[row1, col + 3].Style.Font.Bold = true;
                    workSheet.Cells[row1, 2].Style.Font.Color.SetColor(Cores.CorTexto2);
                    workSheet.Cells[row1, 3 + col].Style.Border.Bottom.Style = ExcelBorderStyle.Thin;
                    workSheet.Cells[row1, 2].Style.Border.Bottom.Style = ExcelBorderStyle.Thin;
                    totLiabilities = row1;
                    row1++;

                }
                row1++;
                support.Divisor("Equity", col, row1, workSheet);
                row1++;
                if (item.PreferredStock != null)
                {
                    support.CommonCaption("Preferred stock", item.PreferredStock, col, row1, workSheet, item);
                    workSheet.Cells[row1, 2].Style.Indent = 2;
                    row1++;
                }
                if (item.CommonStock != null)
                {
                    support.CommonCaption("Common stock", item.CommonStock, col, row1, workSheet, item);
                    workSheet.Cells[row1, 2].Style.Indent = 2;
                    row1++;
                }
                if (item.RetainedEarnings != null)
                {
                    support.CommonCaption("Retained Earnings", item.RetainedEarnings, col, row1, workSheet, item);
                    workSheet.Cells[row1, 2].Style.Indent = 2;
                    row1++;
                }
                if (item.AccumulatedOtherComprehensiveIncomeLoss != null)
                {
                    support.CommonCaption("Accumulated other comprehensive income/loss", item.AccumulatedOtherComprehensiveIncomeLoss, col, row1, workSheet, item);
                    workSheet.Cells[row1, 2].Style.Indent = 2;
                    row1++;
                }
                if (item.OtherTotalStockholdersEquity != null)
                {
                    support.CommonCaption("Other Total stockholders equity", item.OtherTotalStockholdersEquity, col, row1, workSheet, item);
                    workSheet.Cells[row1, 2].Style.Indent = 2;
                    row1++;
                }
                if (item.TotalStockholdersEquity != null)
                {
                    workSheet.Cells[row1, 2].Value = "Total stockholders equity";
                    workSheet.Cells[row1, 2].Style.Font.Bold = true;
                    workSheet.Cells[row1, col + 3].Formula = "=" + colAux + "45+" + colAux + "46+" + colAux + "47+" + colAux + "48+" + colAux + "49"; ;
                    workSheet.Cells[row1, col + 3].Style.Numberformat.Format = "#,##0;(#,##0);-";
                    workSheet.Cells[row1, col + 3].Style.Font.Bold = true;
                   
                    workSheet.Cells[row1, 2].Style.Indent = 2;
                    row1++;
                    row1++;
                }
                if (item.MinorityInterest != null)
                {
                    support.CommonCaption("Minority interest and other non-controlling charges", item.MinorityInterest + (item.TotalAssets -item.TotalLiabilitiesAndTotalEquity), col, row1, workSheet, item);
                    workSheet.Cells[row1, 2].Style.Indent = 2;
                    row1++;
                }
                if (item.TotalEquity != null)
                {
                    workSheet.Cells[row1, 2].Value = "Total stockholders equity";
                    workSheet.Cells[row1, 2].Style.Font.Bold = true;
                    workSheet.Cells[row1, col + 3].Formula = "=" + colAux + "50+" + colAux + "52";
                    workSheet.Cells[row1, col + 3].Style.Numberformat.Format = "#,##0;(#,##0);-";
                    workSheet.Cells[row1, col + 3].Style.Font.Bold = true;
                    
                    workSheet.Cells[row1, 2].Style.Font.Color.SetColor(Cores.CorTexto2);
                    workSheet.Cells[row1, 3 + col].Style.Border.Bottom.Style = ExcelBorderStyle.Thin;
                    workSheet.Cells[row1, 2].Style.Border.Bottom.Style = ExcelBorderStyle.Thin;
                    totEquity = row1;
                    row1++;
                    row1++;
                }


                //acrescentar o check
                workSheet.Cells[row1, 2].Value = "Check";
                workSheet.Cells[row1, 3 + col].Formula = "=" + workSheet.Cells[totAssets, 3 + col].Address + "-(" + workSheet.Cells[totLiabilities, 3 + col].Address + "+" + workSheet.Cells[totEquity, 3 + col].Address + ")";
                workSheet.Cells[row1, 2].Style.Font.Color.SetColor(Color.Red);
                workSheet.Cells[row1, 3 + col].Style.Font.Color.SetColor(Color.Red);
                workSheet.Cells[row1, col + 3].Style.Numberformat.Format = "#,##0;(#,##0);-";

                row1++;
                row1++;
                support.Subtitle("Other captions", col, row1, workSheet);
                row1++;
                if (item.TotalDebt != null)
                {
                    support.CommonCaption("Total debt", item.TotalDebt, col, row1, workSheet, item);
                    workSheet.Cells[row1, 2].Style.Indent = 2;
                    row1++;
                }
                if (item.TotalInvestments != null)
                {
                    support.CommonCaption("Total investments", item.TotalInvestments, col, row1, workSheet, item);
                    workSheet.Cells[row1, 2].Style.Indent = 2;
                    row1++;
                }
                if (item.NetDebt != null)
                {
                    support.CommonCaption("Net debt", item.NetDebt, col, row1, workSheet, item);
                    workSheet.Cells[row1, 2].Style.Indent = 2;
                    workSheet.Cells[row1, 3 + col].Style.Border.Bottom.Style = ExcelBorderStyle.Thin;
                    workSheet.Cells[row1, 2].Style.Border.Bottom.Style = ExcelBorderStyle.Thin;
                    row1++;
                }
                workSheet.Columns[2, 3 + col].AutoFit();
                col++;
            }

            if (quarters > 0)
            {

                for (int i = quarters - 1; i > -1; i--)
                {
                    if (i == 0)
                    {
                        workSheet.Columns[3 + balances.Count()].Width = 3;
                    }
                    int column = 3 + (quarters - i) + balances.Count();
                    int row1 = 5;

                    if (i == quarters - 1)
                    {
                        workSheet.Cells[2, column].Value = "Last Year Quarters ";
                        workSheet.Cells[2, column].Style.Font.Bold = true;
                        workSheet.Cells[2, column].Style.Font.Color.SetColor(Color.White);
                    }
                    workSheet.Cells[2, column].Style.Fill.PatternType = ExcelFillStyle.Solid;
                    workSheet.Cells[2, column].Style.Fill.BackgroundColor.SetColor(Cores.CorPrincipal);



                    for (int a = 0; a < 57; a++)
                    {
                        workSheet.Cells[4 + a, column].Style.Fill.PatternType = ExcelFillStyle.Solid;
                        workSheet.Cells[4 + a, column].Style.Fill.BackgroundColor.SetColor(Cores.CorCinzentoBackground);

                    }

                    ConstructionFinancialStatementsSupport support = new();

                    workSheet.Cells[row1 - 2, column].Value = companyProfile.financialsQuarter.balance[i].Date;
                    workSheet.Cells[row1 - 2, column].Style.Numberformat.Format = DateTimeFormatInfo.CurrentInfo.ShortDatePattern;
                    workSheet.Cells[row1 - 2, column].Style.Font.Bold = true;
                    workSheet.Cells[row1 - 2, column].Style.Font.Color.SetColor(Cores.CorTexto);
                    workSheet.Cells[row1 - 2, column].Style.Border.Bottom.Style = ExcelBorderStyle.Thin;

                    string columnAux = columnName.GetExcelColumnName(column);
                    row1 += 2;

                    //cash 
                    workSheet.Cells[row1, column].Value = companyProfile.financialsQuarter.balance[i].CashAndCashEquivalents / 1000;
                    workSheet.Cells[row1, column].Style.Numberformat.Format = "#,##0;(#,##0);-";
                    row1++;

                    //Short term investments
                    workSheet.Cells[row1, column].Value = companyProfile.financialsQuarter.balance[i].ShortTermInvestments / 1000;
                    workSheet.Cells[row1, column].Style.Numberformat.Format = "#,##0;(#,##0);-";
                    row1++;


                    //accounts receivable
                    workSheet.Cells[row1, column].Value = companyProfile.financialsQuarter.balance[i].AccountsReceivables / 1000;
                    workSheet.Cells[row1, column].Style.Numberformat.Format = "#,##0;(#,##0);-";
                    row1++;

                    //Inventory
                    workSheet.Cells[row1, column].Value = companyProfile.financialsQuarter.balance[i].Inventory/1000;
                    workSheet.Cells[row1, column].Style.Numberformat.Format = "#,##0;(#,##0);-";
                    row1 ++;

                    //Other current assets
                    workSheet.Cells[row1, column].Value = companyProfile.financialsQuarter.balance[i].OtherCurrentAssets / 1000;
                    workSheet.Cells[row1, column].Style.Numberformat.Format = "#,##0;(#,##0);-";
                    row1++;

                    //Total current assets
                    workSheet.Cells[row1, column].Formula = "=" + columnAux + "7+" + columnAux + "8+" + columnAux + "9+" + columnAux + "10+" + columnAux + "11";
                    workSheet.Cells[row1, column].Style.Numberformat.Format = "#,##0;(#,##0);-";
                    workSheet.Cells[row1, column].Style.Font.Bold = true;
                    row1 += 3;

                    //PP&E
                    workSheet.Cells[row1, column].Value = companyProfile.financialsQuarter.balance[i].propertyPlantEquipmentNet / 1000;
                    workSheet.Cells[row1, column].Style.Numberformat.Format = "#,##0;(#,##0);-";
                    row1++;

                    //Goodwill
                    workSheet.Cells[row1, column].Value = companyProfile.financialsQuarter.balance[i].Goodwill / 1000;
                    workSheet.Cells[row1, column].Style.Numberformat.Format = "#,##0;(#,##0);-";
                    row1++;

                    //Intangible assets
                    workSheet.Cells[row1, column].Value = companyProfile.financialsQuarter.balance[i].IntangibleAssets / 1000;
                    workSheet.Cells[row1, column].Style.Numberformat.Format = "#,##0;(#,##0);-";
                    row1++;

                    //Long Term investments
                    workSheet.Cells[row1, column].Value = companyProfile.financialsQuarter.balance[i].LongtermInvestments / 1000;
                    workSheet.Cells[row1, column].Style.Numberformat.Format = "#,##0;(#,##0);-";
                    row1++;

                    //Tax assets
                    workSheet.Cells[row1, column].Value = companyProfile.financialsQuarter.balance[i].TaxAssets / 1000;
                    workSheet.Cells[row1, column].Style.Numberformat.Format = "#,##0;(#,##0);-";
                    row1++;

                    //Other non-current assets
                    workSheet.Cells[row1, column].Value = companyProfile.financialsQuarter.balance[i].OtherNonCurrentAssets / 1000;
                    workSheet.Cells[row1, column].Style.Numberformat.Format = "#,##0;(#,##0);-";
                    row1++;

                    //Total non-current assets
                    workSheet.Cells[row1, column].Formula = "=" + columnAux + "15+" + columnAux + "16+" + columnAux + "17+" + columnAux + "18+" + columnAux + "19+" + columnAux + "20";
                    workSheet.Cells[row1, column].Style.Numberformat.Format = "#,##0;(#,##0);-";
                    workSheet.Cells[row1, column].Style.Font.Bold = true;
                    row1+= 2;

                    //Total assets
                    workSheet.Cells[row1, column].Formula = "=" + columnAux + "12+" + columnAux + "21";
                    workSheet.Cells[row1, column].Style.Numberformat.Format = "#,##0;(#,##0);-";
                    workSheet.Cells[row1, column].Style.Font.Bold = true;
                    workSheet.Cells[row1, column].Style.Border.Bottom.Style = ExcelBorderStyle.Thin;
                    row1 += 4;

                    //accounts payable
                    workSheet.Cells[row1, column].Value = companyProfile.financialsQuarter.balance[i].AccountPayables / 1000;
                    workSheet.Cells[row1, column].Style.Numberformat.Format = "#,##0;(#,##0);-";
                    row1 ++;

                    //short term debt
                    workSheet.Cells[row1, column].Value = companyProfile.financialsQuarter.balance[i].ShortTermDebt / 1000;
                    workSheet.Cells[row1, column].Style.Numberformat.Format = "#,##0;(#,##0);-";
                    row1++;

                    //tax payables
                    //workSheet.Cells[row1, column].Value = companyProfile.financialsQuarter.balance[i].TaxPayables / 1000;
                    workSheet.Cells[row1, column].Style.Numberformat.Format = "#,##0;(#,##0);-";
                    row1++;

                    //Deferred Revenue
                    workSheet.Cells[row1, column].Value = companyProfile.financialsQuarter.balance[i].DeferredRevenue / 1000;
                    workSheet.Cells[row1, column].Style.Numberformat.Format = "#,##0;(#,##0);-";
                    row1++;

                    //Other current liabilities
                    workSheet.Cells[row1, column].Value = companyProfile.financialsQuarter.balance[i].DeferredRevenue / 1000;
                    workSheet.Cells[row1, column].Style.Numberformat.Format = "#,##0;(#,##0);-";
                    row1++;

                    //Total current liabilities
                    workSheet.Cells[row1, column].Formula = "=" + columnAux + "27+" + columnAux + "28+" + columnAux + "29+" + columnAux + "30+" + columnAux + "31";
                    workSheet.Cells[row1, column].Style.Numberformat.Format = "#,##0;(#,##0);-";
                    workSheet.Cells[row1, column].Style.Font.Bold = true;
                    row1+=3;

                    //Long term debt
                    workSheet.Cells[row1, column].Value = companyProfile.financialsQuarter.balance[i].LongTermDebt / 1000;
                    workSheet.Cells[row1, column].Style.Numberformat.Format = "#,##0;(#,##0);-";
                    row1++;

                    //Non current deferred revenue
                    workSheet.Cells[row1, column].Value = companyProfile.financialsQuarter.balance[i].DeferredRevenueNonCurrent / 1000;
                    workSheet.Cells[row1, column].Style.Numberformat.Format = "#,##0;(#,##0);-";
                    row1++;

                    //Deferred tax liabilities
                    workSheet.Cells[row1, column].Value = companyProfile.financialsQuarter.balance[i].DeferredTaxLiabilitiesNonCurrent / 1000;
                    workSheet.Cells[row1, column].Style.Numberformat.Format = "#,##0;(#,##0);-";
                    row1++;

                    //Other liabilities
                    workSheet.Cells[row1, column].Value = companyProfile.financialsQuarter.balance[i].OtherLiabilities / 1000;
                    workSheet.Cells[row1, column].Style.Numberformat.Format = "#,##0;(#,##0);-";
                    row1++;

                    //Other liabilities
                    workSheet.Cells[row1, column].Value = companyProfile.financialsQuarter.balance[i].OtherNonCurrentLiabilities / 1000;
                    workSheet.Cells[row1, column].Style.Numberformat.Format = "#,##0;(#,##0);-";
                    row1++;

                    //Total non-current liabilities
                    workSheet.Cells[row1, column].Formula = "=" + columnAux + "35+" + columnAux + "36+" + columnAux + "37+" + columnAux + "38+" + columnAux + "39";
                    workSheet.Cells[row1, column].Style.Numberformat.Format = "#,##0;(#,##0);-";
                    workSheet.Cells[row1, column].Style.Font.Bold = true;
                    row1+=2;

                    //Total  liabilities
                    workSheet.Cells[row1, column].Formula = "=" + columnAux + "32+" + columnAux + "40";
                    workSheet.Cells[row1, column].Style.Numberformat.Format = "#,##0;(#,##0);-";
                    workSheet.Cells[row1, column].Style.Font.Bold = true;
                    workSheet.Cells[row1, column].Style.Border.Bottom.Style = ExcelBorderStyle.Thin;
                    row1 +=3;

                    //Preferred stock
                    workSheet.Cells[row1, column].Value = companyProfile.financialsQuarter.balance[i].PreferredStock / 1000;
                    workSheet.Cells[row1, column].Style.Numberformat.Format = "#,##0;(#,##0);-";
                    row1 ++;

                    //common stock
                    workSheet.Cells[row1, column].Value = companyProfile.financialsQuarter.balance[i].CommonStock / 1000;
                    workSheet.Cells[row1, column].Style.Numberformat.Format = "#,##0;(#,##0);-";
                    row1++;

                    //Retained Earnings
                    workSheet.Cells[row1, column].Value = companyProfile.financialsQuarter.balance[i].RetainedEarnings / 1000;
                    workSheet.Cells[row1, column].Style.Numberformat.Format = "#,##0;(#,##0);-";
                    row1++;

                    //Accumulated Earnings
                    workSheet.Cells[row1, column].Value = companyProfile.financialsQuarter.balance[i].AccumulatedOtherComprehensiveIncomeLoss / 1000;
                    workSheet.Cells[row1, column].Style.Numberformat.Format = "#,##0;(#,##0);-";
                    row1++;

                    //Other stockholder equity
                    workSheet.Cells[row1, column].Value = companyProfile.financialsQuarter.balance[i].OtherTotalStockholdersEquity / 1000;
                    workSheet.Cells[row1, column].Style.Numberformat.Format = "#,##0;(#,##0);-";
                    row1++;

                    //total stockholder equity
                    workSheet.Cells[row1, column].Formula = "=" + columnAux + "45+" + columnAux + "46+" + columnAux + "47+" + columnAux + "48+" + columnAux + "49";
                    workSheet.Cells[row1, column].Style.Numberformat.Format = "#,##0;(#,##0);-";
                    workSheet.Cells[row1, column].Style.Font.Bold = true;
                    row1+=2;

                    //Minority interest
                    workSheet.Cells[row1, column].Value = companyProfile.financialsQuarter.balance[i].MinorityInterest / 1000;
                    workSheet.Cells[row1, column].Style.Numberformat.Format = "#,##0;(#,##0);-";
                    row1 ++;

                    //Total equity
                    workSheet.Cells[row1, column].Formula = "=" + columnAux + "50+" + columnAux + "52";
                    workSheet.Cells[row1, column].Style.Numberformat.Format = "#,##0;(#,##0);-";
                    workSheet.Cells[row1, column].Style.Border.Bottom.Style = ExcelBorderStyle.Thin;
                    row1++;
                    row1++;


                    //Check
                    workSheet.Cells[row1, column].Formula = "=(" + columnAux + "53+" + columnAux + "42)-" + columnAux + "23";
                    workSheet.Cells[row1, column].Style.Numberformat.Format = "#,##0;(#,##0);-";
                    workSheet.Cells[row1, column].Style.Font.Color.SetColor(Color.Red);
                    row1++;

                    workSheet.Columns[column].AutoFit();

                }
            }


            //Atençao que esta coluna X foi feita ao calhas. temos de fazer isto alterar consoante o input
            //workSheet.Cells["B:X"].AutoFitColumns();
        }
    }
}
