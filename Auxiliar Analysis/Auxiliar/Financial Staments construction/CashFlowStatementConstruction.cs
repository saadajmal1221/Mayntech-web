using Mayntech___Individual_Solution.Auxiliar.Auxiliar.FMP;
using Mayntech___Individual_Solution.Auxiliar.Auxiliar.Other;
using OfficeOpenXml;
using OfficeOpenXml.Style;
using System.Drawing;
using System.Globalization;

namespace Mayntech___Individual_Solution.Auxiliar.Auxiliar
{
    public class CashFlowStatementConstruction
    {
        public int CFO;
        public int CFI;
        public int CFF;

        public async Task CreateBS(ExcelPackage package, List<FinancialStatements> cashFlow, int col, int row, string companyName)
        {
            var workSheet = package.Workbook.Worksheets.Add("CFS");
            workSheet.View.ShowGridLines = false;
            workSheet.View.FreezePanes(4, 3);



            // Cria a coluna azul em cima da tabela
            workSheet.Cells[2, 2].Value = "Cash Flow Statement - " + companyName;
            workSheet.Cells[2, 2].Style.Font.Bold = true;
            workSheet.Cells[2, 2].Style.Font.Color.SetColor(Color.White);

            // atenção que este "H2" está hardcoded. Tem de ser refeito para variar com o número de anos
            workSheet.Cells["B2"].Style.Fill.PatternType = ExcelFillStyle.Solid;
            workSheet.Cells["B2"].Style.Fill.BackgroundColor.SetColor(Cores.CorPrincipal);
            //criar um field com o valor ser em milhoes ou thousands


            cashFlow.Reverse();
            foreach (FinancialStatements item in cashFlow)
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

                if (item.NetIncome != null)
                {
                    support.CommonCaption("Net Income", (double)item.NetIncome, col, row1, workSheet, item);
                    workSheet.Cells[row1, 2].Style.Indent = 2;
                    workSheet.Cells[row1, 2].Style.Font.Bold = true;
                    workSheet.Cells[row1, 3 + col].Style.Font.Bold = true;
                    row1++;
                }
                if (item.DepreciationAndAmortization != null)
                {
                    support.CommonCaption("Depreciation & Amortization", (double)item.DepreciationAndAmortization, col, row1, workSheet, item);
                    workSheet.Cells[row1, 2].Style.Indent = 2;
                    row1++;
                }
                if (item.OtherNonCashItems != null)
                {
                    support.CommonCaption("Other Non Cash Items", item.OtherNonCashItems, col, row1, workSheet, item);
                    workSheet.Cells[row1, 2].Style.Indent = 2;
                    row1++;
                }
                if (item.DeferredIncomeTax != null)
                {
                    support.CommonCaption("Deferred Income tax", item.DeferredIncomeTax, col, row1, workSheet, item);
                    workSheet.Cells[row1, 2].Style.Indent = 2;
                    row1++;
                }
                if (item.StockBasedCompensation != null)
                {
                    support.CommonCaption("Stock Based Compensation", item.StockBasedCompensation, col, row1, workSheet, item);
                    workSheet.Cells[row1, 2].Style.Indent = 2;
                    row1++;
                }
                if (item.ChangeInWorkingCapital != null)
                {
                    support.CommonCaption("Change in Working Capital", item.ChangeInWorkingCapital, col, row1, workSheet, item);
                    workSheet.Cells[row1, 2].Style.Indent = 2;
                    row1++;
                }
                //if (item.AccountsReceivables != null)
                //{
                //    support.CommonSubCaption("Change in Accounts Receivables", -item.AccountsReceivables, col, row1, workSheet, item);
                //    workSheet.Cells[row1, 2].Style.Indent = 4;
                //    row1++;
                //}
                //if (item.AccountsPayables != null)
                //{
                //    support.CommonSubCaption("Change in Accounts Payables", -item.AccountsPayables, col, row1, workSheet, item);
                //    workSheet.Cells[row1, 2].Style.Indent = 4;
                //    row1++;
                //}
                //if (item.Inventory != null)
                //{
                //    support.CommonSubCaption("Change in inventory", -item.Inventory, col, row1, workSheet, item);
                //    workSheet.Cells[row1, 2].Style.Indent = 4;
                //    row1++;
                //}

                //if (item.ChangeInWorkingCapital != null)
                //{
                //    double input = -(item.ChangeInWorkingCapital - item.AccountsPayables - item.AccountsReceivables - item.Inventory - item.OtherWorkingCapital);
                //    support.CommonSubCaption("Changes in other current assets and Liabilities", input, col, row1, workSheet, item);
                //    workSheet.Cells[row1, 2].Style.Indent = 4;
                //    row1++;
                //}
                //if (item.OtherWorkingCapital != null)
                //{
                //    support.CommonSubCaption("Change in Other Working Capital", -item.OtherWorkingCapital, col, row1, workSheet, item);
                //    workSheet.Cells[row1, 2].Style.Indent = 4;
                //    row1++;
                //}

                if (item.NetCashProvidedByOperatingActivities != null)
                {
                    support.CaptionTotal("Cash from Operating Activities", item.NetCashProvidedByOperatingActivities, col, row1, workSheet, item);
                    workSheet.Cells[row1, 2].Style.Font.Color.SetColor(Cores.CorTexto);
                    workSheet.Cells[row1, 3 + col].Style.Font.Color.SetColor(Cores.CorTexto);
                    workSheet.Cells[row1, 3 + col].Style.Border.Bottom.Style = ExcelBorderStyle.Thin;
                    workSheet.Cells[row1, 2].Style.Border.Bottom.Style = ExcelBorderStyle.Thin;
                    CFO = row1;
                    row1++;
                    row1++;


                }
                if (item.InvestmentsInPropertyPlantAndEquipment != null)
                {
                    support.CommonCaption("Investments in PP&E", item.InvestmentsInPropertyPlantAndEquipment, col, row1, workSheet, item);
                    workSheet.Cells[row1, 2].Style.Indent = 2;
                    row1++;
                }
                if (item.AcquisitionsNet != null)
                {
                    support.CommonCaption("Acquisitions, net", item.AcquisitionsNet, col, row1, workSheet, item);
                    workSheet.Cells[row1, 2].Style.Indent = 2;
                    row1++;
                }
                if (item.PurchasesOfInvestments != null)
                {
                    support.CommonCaption("Purchases of investments", item.PurchasesOfInvestments, col, row1, workSheet, item);
                    workSheet.Cells[row1, 2].Style.Indent = 2;
                    row1++;
                }
                if (item.SalesMaturitiesOfInvestments != null)
                {
                    support.CommonCaption("Sales Maturities of Investments", item.SalesMaturitiesOfInvestments, col, row1, workSheet, item);
                    workSheet.Cells[row1, 2].Style.Indent = 2;
                    row1++;
                }
                if (item.otherInvestingActivites != null)
                {
                    support.CommonCaption("Other investing activities", item.otherInvestingActivites, col, row1, workSheet, item);
                    workSheet.Cells[row1, 2].Style.Indent = 2;
                    row1++;
                }
                if (item.netCashUsedForInvestingActivites != null)
                {
                    support.CaptionTotal("Cash from investing Activities", item.netCashUsedForInvestingActivites, col, row1, workSheet, item);
                    workSheet.Cells[row1, 2].Style.Font.Color.SetColor(Cores.CorTexto);
                    workSheet.Cells[row1, 3 + col].Style.Font.Color.SetColor(Cores.CorTexto);
                    workSheet.Cells[row1, 3 + col].Style.Border.Bottom.Style = ExcelBorderStyle.Thin;
                    workSheet.Cells[row1, 2].Style.Border.Bottom.Style = ExcelBorderStyle.Thin;
                    CFI = row1;
                    row1++;
                    row1++;
                }
                if (item.DebtRepayment != null)
                {
                    support.CommonCaption("Debt Repayment", item.DebtRepayment, col, row1, workSheet, item);
                    workSheet.Cells[row1, 2].Style.Indent = 2;
                    row1++;
                }
                if (item.CommonStockIssued != null)
                {
                    support.CommonCaption("Common stock issued", item.CommonStockIssued, col, row1, workSheet, item);
                    workSheet.Cells[row1, 2].Style.Indent = 2;
                    row1++;
                }
                if (item.CommonStockRepurchased != null)
                {
                    support.CommonCaption("Common stock repurchased", item.CommonStockRepurchased, col, row1, workSheet, item);
                    workSheet.Cells[row1, 2].Style.Indent = 2;
                    row1++;
                }
                if (item.DividendsPaid != null)
                {
                    support.CommonCaption("Dividends paid", item.DividendsPaid, col, row1, workSheet, item);
                    workSheet.Cells[row1, 2].Style.Indent = 2;
                    row1++;
                }
                if (item.otherFinancingActivites != null)
                {
                    support.CommonCaption("Other Financing Activities", item.otherFinancingActivites, col, row1, workSheet, item);
                    workSheet.Cells[row1, 2].Style.Indent = 2;
                    row1++;
                }
                if (item.NetCashUsedProvidedByFinancingActivities != null)
                {
                    support.CaptionTotal("Cash from Financing Activities", item.NetCashUsedProvidedByFinancingActivities, col, row1, workSheet, item);
                    workSheet.Cells[row1, 2].Style.Font.Color.SetColor(Cores.CorTexto);
                    workSheet.Cells[row1, 3 + col].Style.Font.Color.SetColor(Cores.CorTexto);
                    workSheet.Cells[row1, 3 + col].Style.Border.Bottom.Style = ExcelBorderStyle.Thin;
                    workSheet.Cells[row1, 2].Style.Border.Bottom.Style = ExcelBorderStyle.Thin;
                    CFF = row1;
                    row1++;
                    row1++;
                }
                if (item.cashAtBeginningOfPeriod != null)
                {
                    support.CommonCaption("Cash at Beginning of period", item.cashAtBeginningOfPeriod, col, row1, workSheet, item);
                    workSheet.Cells[row1, 2].Style.Indent = 2;
                    row1++;
                }
                if (item.NetChangeInCash != null)
                {
                    support.CommonCaption("Change in Cash, net", item.NetChangeInCash, col, row1, workSheet, item);
                    workSheet.Cells[row1, 2].Style.Indent = 2;
                    row1++;
                }
                workSheet.Cells[row1, 2].Value = "Change in cash";
                workSheet.Cells[row1, 3 + col].Formula = "=" + workSheet.Cells[CFO, 3 + col].Address + "+" + workSheet.Cells[CFI, 3 + col].Address + "+" + workSheet.Cells[CFF, 3 + col].Address;
                workSheet.Cells[row1, 2].Style.Fill.PatternType = ExcelFillStyle.Solid;
                workSheet.Cells[row1, 2].Style.Fill.BackgroundColor.SetColor(Cores.corSecundária);
                workSheet.Cells[row1, 3 + col].Style.Fill.PatternType = ExcelFillStyle.Solid;
                workSheet.Cells[row1, 3 + col].Style.Fill.BackgroundColor.SetColor(Cores.corSecundária);
                workSheet.Cells[row1, 3 + col].Style.HorizontalAlignment = ExcelHorizontalAlignment.Right;
                workSheet.Cells[row1, 3 + col].Style.Numberformat.Format = "#,##0;(#,##0);-";
                workSheet.Cells[row1, 2].Style.Indent = 4;
                row1++;

                if (item.EffectOfForexChangesOnCash != null)
                {
                    support.CommonSubCaption("Forex effects", -item.EffectOfForexChangesOnCash, col, row1, workSheet, item);
                    workSheet.Cells[row1, 2].Style.Indent = 4;
                    row1++;

                }




                if (item.CashAtEndOfPeriod != null)
                {
                    support.CaptionTotal("Cash at End of period", item.CashAtEndOfPeriod, col, row1, workSheet, item);
                    workSheet.Cells[row1, 2].Style.Font.Color.SetColor(Cores.CorTexto);
                    workSheet.Cells[row1, 3 + col].Style.Border.Bottom.Style = ExcelBorderStyle.Thin;
                    workSheet.Cells[row1, 2].Style.Border.Bottom.Style = ExcelBorderStyle.Thin;
                    row1++;
                    row1++;
                }
                support.Subtitle("Other captions", col, row1, workSheet);
                row1++;

                if (item.OperatingCashFlow != null)
                {
                    support.CommonCaption("Operating Cash flow", item.OperatingCashFlow, col, row1, workSheet, item);
                    workSheet.Cells[row1, 2].Style.Indent = 2;
                    row1++;
                }
                if (item.CapitalExpenditure != null)
                {
                    support.CommonCaption("Capital expenditure", item.CapitalExpenditure, col, row1, workSheet, item);
                    workSheet.Cells[row1, 2].Style.Indent = 2;
                    row1++;
                }
                if (item.FreeCashFlow != null)
                {
                    support.CommonCaption("Free Cash flow", item.FreeCashFlow, col, row1, workSheet, item);
                    workSheet.Cells[row1, 2].Style.Indent = 2;
                    row1++;
                }

                workSheet.Columns[2, 3 + col].AutoFit();
                col++;
            }
            //Atençao que esta coluna X foi feita ao calhas. temos de fazer isto alterar consoante o input
            //workSheet.Cells["B:X"].AutoFitColumns();
        }
    }
}
