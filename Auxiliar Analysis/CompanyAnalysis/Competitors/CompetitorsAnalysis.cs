using Mayntech___Individual_Solution.Auxiliar.Auxiliar.FMP;
using Mayntech___Individual_Solution.Auxiliar.Auxiliar.Other;
using Mayntech___Individual_Solution.Pages;
using OfficeOpenXml;
using OfficeOpenXml.Style;
using System.Drawing;
using System.Reflection.Metadata.Ecma335;

namespace Mayntech___Individual_Solution.Auxiliar.CompanyAnalysis.Competitors
{
    public class CompetitorsAnalysis : CompetitorsAux
    {
        List<List<string>> addExecutiveSummary = new List<List<string>>();
        List<double> RevenueReferencePage = new List<double>();
        List<double> RevenuePeersPage = new List<double>();
        public void IncomeStatementCompetitors(ExcelPackage package, string companyName, int numberOfYears)
        {
            ExcelNextCol columnName = new ExcelNextCol();
            var worksheet = package.Workbook.Worksheets.Add("P&L - Analysis Competitors");
            worksheet.View.ShowGridLines = false;
            worksheet.View.ZoomScale = 80;
            worksheet.View.FreezePanes(4, 3);
            int row = 2;
            int col = 2;
            worksheet.Cells[row, col].Value = "P&L - Analysis Competitors";
            worksheet.Cells[row, col].Style.Font.Bold = true;
            worksheet.Cells[row, col].Style.Font.Color.SetColor(Color.White);

            int numberOfColumns = Math.Min(5, numberOfYears);

            for (int i = 0; i < numberOfColumns+4; i++)
            {
                if (i!=numberOfColumns+1)
                {
                    worksheet.Cells[row, col + i].Style.Fill.PatternType = ExcelFillStyle.Solid;
                    worksheet.Cells[row, col + i].Style.Fill.BackgroundColor.SetColor(Cores.CorPrincipal);
                }


                if (i==0)
                {
                    worksheet.Cells[row + +1, col].Value = "Description (in '000 " + SolutionModel.incomeStatement[SolutionModel.incomeStatement.Count() - 1].ReportedCurrency + ")";
                    worksheet.Cells[row + 1, col].Style.Font.Bold = true;
                    worksheet.Cells[row + 1, col].Style.Font.Color.SetColor(Cores.CorTexto);
                    worksheet.Cells[row + 1, col].Style.Border.Bottom.Style = ExcelBorderStyle.Thin;
                }
                else if (i>0 && i <numberOfColumns+1)
                {
                    int aux = SolutionModel.NumberYears - (numberOfColumns-1) + 1 + i;
                    string column = columnName.GetExcelColumnName(aux);
                    worksheet.Cells[row + 1, col + i].Formula = "=YEAR('P&L'!" + column + "3)";
                    worksheet.Cells[row + 1, col + i].Style.HorizontalAlignment = ExcelHorizontalAlignment.Right;
                    worksheet.Cells[row + 1, col + i].Style.Font.Color.SetColor(Cores.CorTexto);
                    worksheet.Cells[row + 1, col + i].Style.Border.Bottom.Style = ExcelBorderStyle.Thin;
                    worksheet.Cells[row + 1, col + i].Style.Font.Bold = true;
                }
                else if (i== numberOfColumns + 2)
                {
                    worksheet.Cells[row + 1, col + i].Value = "CAGR";
                    worksheet.Cells[row + 1, col + i].Style.HorizontalAlignment = ExcelHorizontalAlignment.Right;
                    worksheet.Cells[row + 1, col + i].Style.Font.Color.SetColor(Cores.CorTexto);
                    worksheet.Cells[row + 1, col + i].Style.Border.Bottom.Style = ExcelBorderStyle.Thin;
                    worksheet.Cells[row + 1, col + i].Style.Font.Bold = true;
                }
                else if (i == numberOfColumns + 3)
                {
                    worksheet.Cells[row + 1, col + i].Value = "Coefficient of Variation";
                    worksheet.Cells[row + 1, col + i].Style.HorizontalAlignment = ExcelHorizontalAlignment.Right;
                    worksheet.Cells[row + 1, col + i].Style.Font.Color.SetColor(Cores.CorTexto);
                    worksheet.Cells[row + 1, col + i].Style.Border.Bottom.Style = ExcelBorderStyle.Thin;
                    worksheet.Cells[row + 1, col + i].Style.Font.Bold = true;
                }

            }
            row += 2;




            List<double> RevenuePeers = new List<double>();
            List<double> GrossMarginPeers = new List<double>();
            List<double> OperatingMarginPeers = new List<double>();
            List<double> EbitdaMarginPeers = new List<double>();
            List<double> NetIncomeMarginPeers = new List<double>();

            List<double> RevenueRef = new List<double>();
            List<double> GrossMarginRef = new List<double>();
            List<double> OperatingMarginRef = new List<double>();
            List<double> EbitdaMarginRef = new List<double>();
            List<double> NetIncomeMarginRef = new List<double>();

            for (int a = 0; a < numberOfColumns; a++)
            {
                try
                {
                    List<double> revenueAux = new List<double>();
                    List<double> GrossMarginAux = new List<double>();
                    List<double> OperatingMarginAux = new List<double>();
                    List<double> EbitdaAux = new List<double>();
                    List<double> NetIncomeAux = new List<double>();

                    RevenueRef.Add((double)SolutionModel.incomeStatement[numberOfColumns - a-1].Revenue);
                    int yearRef = SolutionModel.incomeStatement[numberOfColumns - a - 1].Date.Year;
                    int month = SolutionModel.incomeStatement[numberOfColumns - a - 1].Date.Month;
                    GrossMarginRef.Add((double)SolutionModel.incomeStatement[numberOfColumns - a - 1].GrossProfitRatio);
                    OperatingMarginRef.Add((double)SolutionModel.incomeStatement[numberOfColumns - a - 1].OperatingIncomeRatio);
                    EbitdaMarginRef.Add((double)SolutionModel.incomeStatement[numberOfColumns - a - 1].EBITDARatio);
                    NetIncomeMarginRef.Add((double)SolutionModel.incomeStatement[numberOfColumns - a - 1].NetIncomeRatio);

                    foreach (KeyValuePair<string, List<FinancialStatements>> item in SolutionModel.IncomeStatementDict)
                    {
                        try
                        {
                            if (item.Value.Count() >= Math.Min(numberOfYears, 4))
                            {
                                revenueAux.Add((double)item.Value[a].Revenue);
                                GrossMarginAux.Add((double)item.Value[a].GrossProfitRatio);
                                OperatingMarginAux.Add((double)item.Value[a].OperatingIncomeRatio);
                                EbitdaAux.Add((double)item.Value[a].EBITDARatio);
                                NetIncomeAux.Add((double)item.Value[a].NetIncomeRatio);
                            }
                        }
                        catch (Exception)
                        {

                            continue;
                        }


                    }

                    RevenuePeers.Add(revenueAux.Average());
                    GrossMarginPeers.Add(GrossMarginAux.Average());
                    OperatingMarginPeers.Add(OperatingMarginAux.Average());
                    EbitdaMarginPeers.Add(EbitdaAux.Average());
                    NetIncomeMarginPeers.Add(NetIncomeAux.Average());
                }
                catch (Exception)
                {
                   
                }

            }
            if (RevenuePeers.Count()<RevenueRef.Count())
            {
                RevenueRef.RemoveRange(RevenueRef.Count()-2,1);
                GrossMarginRef.RemoveAt(RevenueRef.Count() - 2);
                OperatingMarginRef.RemoveAt(RevenueRef.Count() - 2);
                EbitdaMarginRef.RemoveAt(RevenueRef.Count() - 2);
                NetIncomeMarginRef.RemoveAt(RevenueRef.Count() - 2);
            }
            RevenueReferencePage = RevenueRef;
            RevenuePeers.Reverse();
            RevenuePeersPage = RevenuePeers;
            GrossMarginPeers.Reverse();
            EbitdaMarginPeers.Reverse();
            NetIncomeMarginPeers.Reverse();

                string revEvolution = Evolution(RevenueRef, RevenuePeers);
                string RevSize = Size(RevenueRef, RevenuePeers);
                string CoeficcientRev = coefficientOfVariation(RevenueRef, RevenuePeers);

                string GrossEvolution = Evolution(GrossMarginRef, GrossMarginPeers);
                string GrossSize = Size(GrossMarginRef, GrossMarginPeers);
                string CoeficcientGross = coefficientOfVariation(GrossMarginRef, GrossMarginPeers);

                string OperatingEvolution = Evolution(OperatingMarginRef, OperatingMarginPeers);
                string OperatingSize = Size(OperatingMarginRef, OperatingMarginPeers);
                string CoeficcientOperating = coefficientOfVariation(OperatingMarginRef, OperatingMarginPeers);

                string EbitdaEvolution = Evolution(EbitdaMarginRef, EbitdaMarginPeers);
                string EbitdaSize = Size(EbitdaMarginRef, EbitdaMarginPeers);
                string CoeficcientEbitda = coefficientOfVariation(EbitdaMarginRef, EbitdaMarginPeers);

                string NetIncomeEvolution = Evolution(NetIncomeMarginRef, NetIncomeMarginPeers);
                string NetIncomeSize = Size(NetIncomeMarginRef, NetIncomeMarginPeers);
                string CoeficcientNEtIncome = coefficientOfVariation(NetIncomeMarginRef, NetIncomeMarginPeers);



            List<string> commentList = new List<string> { "Growing at a much faster pace than its peers. ",
            "Growing faster than its peers. ", 
                "Growing at same pace as its peers. ",
            "Decreasing more than competitors. What is the reason behind this decrease? ",
            "Decreasing at a much faster pace than its peers. What is the reason behind this decrease?",
            "Decreasing less than competitors. ",
            "Increasing less than competitors. ",
            "Increasing, unlike its competitors. ",
            "Decreasing, unlike its competitors. "};

            List<string> commentListCoefficient = new List<string> { "Very low volatility in relation to competitors. ", "Low volatility in relation to competitors. ",
            "Same volatility as its competitors. ", "High volatility in relation to competitors. ",
                "Very high volatility in relation to competitors. "};

            List<string> commentListSize = new List<string> { "Much higher margin than competitors. ", "Higher margin than competitors. ",
            "Same margin as competitors. ", "Smaller margin than competitors. ",
                "Much smaller margin than competitors. "};

            List<int> auxiliar = new List<int>();


            List<List<string>> listRevenue = SupportCompetitors(worksheet, row, col, RevenuePeers, RevenueRef, revEvolution, RevSize, 5, "Revenue", companyName, "+", commentList, auxiliar, CoeficcientRev, commentListSize, commentListCoefficient, numberOfColumns);
            for (int i = 0; i < listRevenue.Count(); i++)
            {
                addExecutiveSummary.Add(listRevenue[i]);
            }
            row += 6;

            List<List<string>> listGrossMargin = SupportCompetitors(worksheet, row, col, GrossMarginPeers, GrossMarginRef, GrossEvolution, GrossSize, 8, "Gross Margin", companyName, "+", commentList, auxiliar, CoeficcientGross, commentListSize, commentListCoefficient, numberOfColumns);
            for (int i = 0; i < listGrossMargin.Count(); i++)
            {
                addExecutiveSummary.Add(listGrossMargin[i]);
            }
            row += 6;

            List<List<string>> listOperating = SupportCompetitors(worksheet, row, col, OperatingMarginPeers, OperatingMarginRef, OperatingEvolution, OperatingSize, 16, "Operating income Margin", companyName, "+", commentList, auxiliar, CoeficcientOperating, commentListSize, commentListCoefficient, numberOfColumns);
            for (int i = 0; i < listOperating.Count(); i++)
            {
                addExecutiveSummary.Add(listOperating[i]);
            }
            row += 6;

            List<List<string>> listEbitda = SupportCompetitors(worksheet, row, col, EbitdaMarginPeers, EbitdaMarginRef, EbitdaEvolution, EbitdaSize, 31, "EBITDA Margin", companyName, "+", commentList, auxiliar, CoeficcientEbitda, commentListSize, commentListCoefficient, numberOfColumns);
            for (int i = 0; i < listEbitda.Count(); i++)
            {
                addExecutiveSummary.Add(listEbitda[i]);
            }
            row += 6;

            //List<List<string>> listNetIncome = SupportCompetitors(worksheet, row, col, NetIncomeMarginPeers, NetIncomeMarginRef, NetIncomeEvolution, NetIncomeSize, 24, "Net Income Margin", companyName, "+", commentList, auxiliar, CoeficcientNEtIncome, commentListSize, commentListCoefficient);
            //for (int i = 0; i < listNetIncome.Count(); i++)
            //{
            //    addExecutiveSummary.Add(listNetIncome[i]);
            //}
            //row += 6;

            ExecutiveSummaryFromIncomestatementAnalysis(addExecutiveSummary);
            

        }

        public void BalanceSheetCompetitors(ExcelPackage package, string companyName, int numberOfYears)
        {
            ExcelNextCol columnName = new ExcelNextCol();
            var worksheet = package.Workbook.Worksheets.Add("BS - Analysis Competitors");
            worksheet.View.ShowGridLines = false;
            worksheet.View.ZoomScale = 80;
            worksheet.View.FreezePanes(4, 3);
            int row = 2;
            int col = 2;
            worksheet.Cells[row, col].Value = "BS - Analysis Competitors";
            worksheet.Cells[row, col].Style.Font.Bold = true;
            worksheet.Cells[row, col].Style.Font.Color.SetColor(Color.White);

            int numberOfColumns = Math.Min(5, numberOfYears);

            for (int i = 0; i < numberOfColumns + 4; i++)
            {
                if (i != 6)
                {
                    worksheet.Cells[row, col + i].Style.Fill.PatternType = ExcelFillStyle.Solid;
                    worksheet.Cells[row, col + i].Style.Fill.BackgroundColor.SetColor(Cores.CorPrincipal);
                }

                if (i == 0)
                {
                    worksheet.Cells[row + +1, col].Value = "Description (in '000 " + SolutionModel.incomeStatement[SolutionModel.incomeStatement.Count() - 1].ReportedCurrency + ")";
                    worksheet.Cells[row + 1, col].Style.Font.Bold = true;
                    worksheet.Cells[row + 1, col].Style.Font.Color.SetColor(Cores.CorTexto);
                    worksheet.Cells[row + 1, col].Style.Border.Bottom.Style = ExcelBorderStyle.Thin;
                }
                else if (i > 0 && i < 6)
                {
                    int aux = SolutionModel.NumberYears - 4 + 1 + i;
                    string column = columnName.GetExcelColumnName(aux);
                    worksheet.Cells[row + 1, col + i].Formula = "=YEAR('P&L'!" + column + "3)";
                    worksheet.Cells[row + 1, col + i].Style.HorizontalAlignment = ExcelHorizontalAlignment.Right;
                    worksheet.Cells[row + 1, col + i].Style.Font.Color.SetColor(Cores.CorTexto);
                    worksheet.Cells[row + 1, col + i].Style.Border.Bottom.Style = ExcelBorderStyle.Thin;
                    worksheet.Cells[row + 1, col + i].Style.Font.Bold = true;
                }
                else if (i == 7)
                {
                    worksheet.Cells[row + 1, col + i].Value = "CAGR";
                    worksheet.Cells[row + 1, col + i].Style.HorizontalAlignment = ExcelHorizontalAlignment.Right;
                    worksheet.Cells[row + 1, col + i].Style.Font.Color.SetColor(Cores.CorTexto);
                    worksheet.Cells[row + 1, col + i].Style.Border.Bottom.Style = ExcelBorderStyle.Thin;
                    worksheet.Cells[row + 1, col + i].Style.Font.Bold = true;
                }
                else if (i == 8)
                {
                    worksheet.Cells[row + 1, col + i].Value = "Coefficient of Variation";
                    worksheet.Cells[row + 1, col + i].Style.HorizontalAlignment = ExcelHorizontalAlignment.Right;
                    worksheet.Cells[row + 1, col + i].Style.Font.Color.SetColor(Cores.CorTexto);
                    worksheet.Cells[row + 1, col + i].Style.Border.Bottom.Style = ExcelBorderStyle.Thin;
                    worksheet.Cells[row + 1, col + i].Style.Font.Bold = true;
                }

            }
            row += 2;

            List<double> accountsReceivableRef = new List<double>();
            List<double> InventoriesRef = new List<double>();
            List<double> currentAssetsRef = new List<double>();
            List<double> PropertyPlantEquipmRef = new List<double>();
            List<double> LongTermInvestmentsRef = new List<double>();
            List<double> IntangibleAssetsRef = new List<double>();
            List<double> GoodwillRef = new List<double>();
            List<double> AccountsPayableRef = new List<double>();
            List<double> OtherLiabilitiesRef = new List<double>();
            List<double> ShorTermDebtRef = new List<double>();
            List<double> LongtermDebtRef = new List<double>();
            List<double> EquityRef = new List<double>();


            List<double> accountsReceivablePeers = new List<double>();
            List<double> InventoriesPeers = new List<double>();
            List<double> currentAssetsPeers = new List<double>();
            List<double> PropertyPlantEquipmPeers = new List<double>();
            List<double> LongTermInvestmentsPeers = new List<double>();
            List<double> intangbleAssetPeers = new List<double>();
            List<double> goodwillPeers = new List<double>();
            List<double> AccountsPayablePeers = new List<double>();
            List<double> otherLiabilitiesPeers = new List<double>();
            List<double> ShorTermDebtPeers = new List<double>();
            List<double> LongtermDebtPeers = new List<double>();
            List<double> EquityPeers = new List<double>();

            for (int a = 0; a < 5; a++)
            {
                List<double> receivablesAux = new List<double>();
                List<double> inventoriesAux = new List<double>();
                List<double> currentassetsAux = new List<double>();
                List<double> PPEAux = new List<double>();
                List<double> LongTerminvestmentsAux = new List<double>();
                List<double> IntangibleassetsAux = new List<double>();
                List<double> otherLiabilitiesAux = new List<double>();
                List<double> GoodwillAux = new List<double>();
                List<double> payablesAux = new List<double>();
                List<double> shorTermdebtAux = new List<double>();
                List<double> longTermdebtAux = new List<double>();
                List<double> equityAux = new List<double>();

                try
                {
                    accountsReceivableRef.Add(SolutionModel.balances[numberOfColumns-a-1].NetReceivables / SolutionModel.balances[numberOfColumns - a - 1].TotalAssets);
                    InventoriesRef.Add(SolutionModel.balances[numberOfColumns - a - 1].Inventory / SolutionModel.balances[numberOfColumns - a - 1].TotalAssets);
                    currentAssetsRef.Add(SolutionModel.balances[numberOfColumns - a - 1].TotalCurrentAssets / SolutionModel.balances[numberOfColumns - a - 1].TotalAssets);
                    PropertyPlantEquipmRef.Add(SolutionModel.balances[numberOfColumns - a - 1].propertyPlantEquipmentNet / SolutionModel.balances[numberOfColumns - a - 1].TotalAssets);
                    LongTermInvestmentsRef.Add(SolutionModel.balances[numberOfColumns - a - 1].LongtermInvestments / SolutionModel.balances[numberOfColumns - a - 1].TotalAssets);
                    IntangibleAssetsRef.Add(SolutionModel.balances[numberOfColumns - a - 1].IntangibleAssets / SolutionModel.balances[numberOfColumns - a - 1].TotalAssets);
                    GoodwillRef.Add(SolutionModel.balances[numberOfColumns - a - 1].Goodwill / SolutionModel.balances[numberOfColumns - a - 1].TotalAssets);
                    AccountsPayableRef.Add(SolutionModel.balances[numberOfColumns - a - 1].AccountPayables / SolutionModel.balances[numberOfColumns - a - 1].TotalAssets);
                    OtherLiabilitiesRef.Add((SolutionModel.balances[numberOfColumns - a - 1].OtherCurrentLiabilities + SolutionModel.balances[numberOfColumns - a - 1].OtherNonCurrentLiabilities) / SolutionModel.balances[numberOfColumns - a - 1].TotalAssets);
                    EquityRef.Add(SolutionModel.balances[numberOfColumns - a - 1].TotalEquity / SolutionModel.balances[numberOfColumns - a - 1].TotalAssets);


                    foreach (KeyValuePair<string, List<FinancialStatements>> item in SolutionModel.BalanceSheetDict)
                    {
                        try
                        {
                            if (item.Value.Count() >= Math.Min(numberOfYears, 4));
                            {
                                receivablesAux.Add(item.Value[a].NetReceivables / item.Value[a].TotalAssets);
                                inventoriesAux.Add(item.Value[a].Inventory / item.Value[a].TotalAssets);
                                currentassetsAux.Add(item.Value[a].TotalCurrentAssets / item.Value[a].TotalAssets);
                                PPEAux.Add(item.Value[a].propertyPlantEquipmentNet / item.Value[a].TotalAssets);
                                LongTerminvestmentsAux.Add(item.Value[a].LongtermInvestments / item.Value[a].TotalAssets);
                                payablesAux.Add(item.Value[a].AccountPayables / item.Value[a].TotalAssets);
                                equityAux.Add(item.Value[a].TotalEquity / item.Value[a].TotalAssets);
                                GoodwillAux.Add(item.Value[a].Goodwill / item.Value[a].TotalAssets);
                                IntangibleassetsAux.Add(item.Value[a].IntangibleAssets / item.Value[a].TotalAssets);
                                otherLiabilitiesAux.Add((item.Value[a].OtherCurrentLiabilities + item.Value[a].OtherNonCurrentLiabilities) / item.Value[a].TotalAssets);
                            }
                        }
                        catch (Exception)
                        {

                            continue;
                        }

                    }

                    accountsReceivablePeers.Add(receivablesAux.Average());
                    InventoriesPeers.Add(inventoriesAux.Average());
                    currentAssetsPeers.Add(currentassetsAux.Average());
                    PropertyPlantEquipmPeers.Add(PPEAux.Average());
                    LongTermInvestmentsPeers.Add(LongTerminvestmentsAux.Average());
                    EquityPeers.Add(equityAux.Average());
                    AccountsPayablePeers.Add(payablesAux.Average());
                    goodwillPeers.Add(GoodwillAux.Average());
                    intangbleAssetPeers.Add(IntangibleassetsAux.Average());
                    otherLiabilitiesPeers.Add(otherLiabilitiesAux.Average());
                }
                catch (Exception)
                {

                    
                }

            }

            accountsReceivablePeers.Reverse();
            InventoriesPeers.Reverse();
            currentAssetsPeers.Reverse();
            PropertyPlantEquipmPeers.Reverse();
            LongTermInvestmentsPeers.Reverse();
            EquityPeers.Reverse();
            AccountsPayablePeers.Reverse();
            goodwillPeers.Reverse();
            intangbleAssetPeers.Reverse();
            otherLiabilitiesPeers.Reverse();

            if (accountsReceivableRef.Count()>accountsReceivablePeers.Count())
            {
                accountsReceivableRef.RemoveAt(accountsReceivableRef.Count() - 1);
                InventoriesRef.RemoveAt(accountsReceivableRef.Count() - 1);
                currentAssetsRef.RemoveAt(accountsReceivableRef.Count() - 1);
                PropertyPlantEquipmRef.RemoveAt(accountsReceivableRef.Count() - 1);
                LongTermInvestmentsRef.RemoveAt(accountsReceivableRef.Count() - 1);
                IntangibleAssetsRef.RemoveAt(accountsReceivableRef.Count() - 1);
                GoodwillRef.RemoveAt(accountsReceivableRef.Count() - 1);
                AccountsPayableRef.RemoveAt(accountsReceivableRef.Count() - 1);
                OtherLiabilitiesRef.RemoveAt(accountsReceivableRef.Count() - 1);
                EquityRef.RemoveAt(accountsReceivableRef.Count() - 1);

            }



            string receivablesEvolution = EvolutionBS(accountsReceivableRef, accountsReceivablePeers);
            string receivablesSize = Size(accountsReceivableRef, accountsReceivablePeers);
            string receivablesCoeficcient = coefficientOfVariation(accountsReceivableRef, accountsReceivablePeers);

            string InventEvolution = EvolutionBS(InventoriesRef, InventoriesPeers);
            string InventSize = Size(InventoriesRef, InventoriesPeers);
            string InventCoeficcient = coefficientOfVariation(InventoriesRef, InventoriesPeers);

            string CurrAsseEvolution = EvolutionBS(currentAssetsRef, currentAssetsPeers);
            string CurrAsseSize = Size(currentAssetsRef, currentAssetsPeers);
            string CurrAsseCoeficcient = coefficientOfVariation(currentAssetsRef, currentAssetsPeers);

            string PPEEvolution = EvolutionBS(PropertyPlantEquipmRef, PropertyPlantEquipmPeers);
            string PPESize = Size(PropertyPlantEquipmRef, PropertyPlantEquipmPeers);
            string PPECoeficcient = coefficientOfVariation(PropertyPlantEquipmRef, PropertyPlantEquipmPeers);

            string LongTermInvEvolution = EvolutionBS(LongTermInvestmentsRef, LongTermInvestmentsPeers);
            string LongTermInvSize = Size(LongTermInvestmentsRef, LongTermInvestmentsPeers);
            string LongTermInvCoeficcient = coefficientOfVariation(LongTermInvestmentsRef, LongTermInvestmentsPeers);

            string equityEvolution = EvolutionBS(EquityRef, EquityPeers);
            string equitySize = Size(EquityRef, EquityPeers);
            string equityCoeficcient = coefficientOfVariation(EquityRef, EquityPeers);

            string payablesEvolution = EvolutionBS(AccountsPayableRef, AccountsPayablePeers);
            string payablesSize = Size(AccountsPayableRef, AccountsPayablePeers);
            string payablesCoeficcient = coefficientOfVariation(AccountsPayableRef, AccountsPayablePeers);

            string goodwillEvolution = EvolutionBS(GoodwillRef, goodwillPeers);
            string goodwillSize = Size(GoodwillRef, goodwillPeers);
            string goodwillCoeficcient = coefficientOfVariation(GoodwillRef, goodwillPeers);

            string intangibleEvolution = EvolutionBS(IntangibleAssetsRef, intangbleAssetPeers);
            string intangibleSize = Size(IntangibleAssetsRef, intangbleAssetPeers);
            string intangibleCoeficcient = coefficientOfVariation(IntangibleAssetsRef, intangbleAssetPeers);


            string OtherLiabiEvolution = EvolutionBS(OtherLiabilitiesRef, otherLiabilitiesPeers);
            string OtherLiabiSize = Size(OtherLiabilitiesRef, otherLiabilitiesPeers);
            string otherLiabCoeficcient = coefficientOfVariation(OtherLiabilitiesRef, otherLiabilitiesPeers);

            List<string> commentList = new List<string> { "Growing at a much faster pace than its peers. ",
            "Growing faster than its peers. ",
                "Growing at same pace as its peers. ",
            "Decreasing more than competitors. ",
            "Decreasing at a much faster pace than its peers. ",
            "Decreasing less than competitors. ",
            "Increasing less than competitors. ",
            "Increasing, unlike its competitors. ",
            "Decreasing, unlike its competitors. "};

            List<string> commentListCoefficient = new List<string> { "Very low volatility in relation to competitors. ", "Low volatility in relation to competitors. ",
            "Same volatility as its competitors. ", "High volatility in relation to competitors. ",
                "Very high volatility in relation to competitors. "};

            List<string> commentListSize = new List<string> { "Much higher margin than competitors. ", "Higher margin than competitors. ",
            "Same margin as competitors. ", "Smaller margin than competitors. ",
                "Much smaller margin than competitors. "};

            List<int> auxiliar = new List<int>();

            List<List<string>> listreceivables = SupportCompetitors(worksheet, row, col, accountsReceivablePeers, accountsReceivableRef, receivablesEvolution, receivablesSize, 9, "Accounts Receivable", companyName, "+", commentList, auxiliar, receivablesCoeficcient, commentListSize, commentListCoefficient, numberOfColumns);
            for (int i = 0; i < listreceivables.Count(); i++)
            {
                addExecutiveSummary.Add(listreceivables[i]);
            }
            row += 6;

            List<List<string>> listinventory = SupportCompetitors(worksheet, row, col, InventoriesPeers, InventoriesRef, InventEvolution, InventSize, 10, "Inventory", companyName, "+", commentList, auxiliar, InventCoeficcient, commentListSize, commentListCoefficient, numberOfColumns);
            for (int i = 0; i < listinventory.Count(); i++)
            {
                addExecutiveSummary.Add(listinventory[i]);
            }
            row += 6;

            List<List<string>> listcurrentAsse = SupportCompetitors(worksheet, row, col, currentAssetsPeers, currentAssetsRef, CurrAsseEvolution, CurrAsseSize, 12, "Current Assets", companyName, "+", commentList, auxiliar, CurrAsseCoeficcient, commentListSize, commentListCoefficient, numberOfColumns);
            for (int i = 0; i < listcurrentAsse.Count(); i++)
            {
                addExecutiveSummary.Add(listcurrentAsse[i]);
            }
            row += 6;

            List<List<string>> listPPE = SupportCompetitors(worksheet, row, col, PropertyPlantEquipmPeers, PropertyPlantEquipmRef, PPEEvolution, PPESize, 15, "Property Plant & Equipment", companyName, "+", commentList, auxiliar, PPECoeficcient, commentListSize, commentListCoefficient, numberOfColumns);
            for (int i = 0; i < listPPE.Count(); i++)
            {
                addExecutiveSummary.Add(listPPE[i]);
            }
            row += 6;

            List<List<string>> listgoodwill = SupportCompetitors(worksheet, row, col, goodwillPeers, GoodwillRef, goodwillEvolution, goodwillSize, 16, "Goodwill", companyName, "+", commentList, auxiliar, goodwillCoeficcient, commentListSize, commentListCoefficient, numberOfColumns);
            for (int i = 0; i < listgoodwill.Count(); i++)
            {
                addExecutiveSummary.Add(listgoodwill[i]);
            }
            row += 6;

            List<List<string>> listIntangible = SupportCompetitors(worksheet, row, col, intangbleAssetPeers, IntangibleAssetsRef, intangibleEvolution, intangibleSize, 17, "Intangible Assets", companyName, "+", commentList, auxiliar, intangibleCoeficcient, commentListSize, commentListCoefficient, numberOfColumns);
            for (int i = 0; i < listIntangible.Count(); i++)
            {
                addExecutiveSummary.Add(listIntangible[i]);
            }
            row += 6;

            List<List<string>> ListLongInvestments = SupportCompetitors(worksheet, row, col, LongTermInvestmentsPeers, LongTermInvestmentsRef, LongTermInvEvolution, LongTermInvSize, 18, "Long Term Investments", companyName, "+", commentList, auxiliar, LongTermInvCoeficcient, commentListSize, commentListCoefficient, numberOfColumns);
            for (int i = 0; i < ListLongInvestments.Count(); i++)
            {
                addExecutiveSummary.Add(ListLongInvestments[i]);
            }
            row += 6;

            List<List<string>> Listpayables = SupportCompetitors(worksheet, row, col, AccountsPayablePeers, AccountsPayableRef, payablesEvolution, payablesSize, 27, "Accounts Payable", companyName, "-", commentList, auxiliar, payablesCoeficcient, commentListSize, commentListCoefficient, numberOfColumns);
            for (int i = 0; i < Listpayables.Count(); i++)
            {
                addExecutiveSummary.Add(Listpayables[i]);
            }
            row += 6;

            List<List<string>> ListOtherLiab = SupportCompetitors(worksheet, row, col, otherLiabilitiesPeers, OtherLiabilitiesRef, OtherLiabiEvolution, OtherLiabiSize, 31, "Other Liabilities", companyName, "-", commentList, auxiliar, otherLiabCoeficcient, commentListSize, commentListCoefficient, numberOfColumns);
            for (int i = 0; i < ListOtherLiab.Count(); i++)
            {
                addExecutiveSummary.Add(ListOtherLiab[i]);
            }



            row += 6;

        }
    }
}
