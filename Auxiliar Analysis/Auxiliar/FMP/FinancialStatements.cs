namespace Mayntech___Individual_Solution.Auxiliar.Auxiliar.FMP
{
    public class FinancialStatements
    {
        public DateTime Date { get; set; }
        public string? Symbol { get; set; }
        public string? ReportedCurrency { get; set; }

        public string? Cik { get; set; }

        public string? FillingDate { get; set; }

        public double Revenue { get; set; }

        public double CostOfRevenue { get; set; }

        public double GrossProfit { get; set; }

        public double GrossProfitRatio { get; set; }

        public double ResearchAndDevelopmentExpenses { get; set; }

        public double GeneralAndAdministrativeExpenses { get; set; }

        public double SellingAndMarketingExpenses { get; set; }

        public double SellingGeneralAndAdministrativeExpenses { get; set; }

        public double otherExpenses { get; set; }

        public double? interestIncome { get; set; }

        public double OperatingExpenses { get; set; }
        public double CostAndExpenses { get; set; }

        public double? InterestExpense { get; set; }
        public double? DepreciationAndAmortization { get; set; }

        public double EBITDA { get; set; }

        public double EBITDARatio { get; set; }
        public double OperatingIncome { get; set; }
        public double OperatingIncomeRatio { get; set; }
        public double TotalOtherIncomeExpensesNet { get; set; }

        public double? IncomeBeforeTax { get; set; }
        public double IncomeBeforeTaxRatio { get; set; }

        public double? IncomeTaxExpense { get; set; }

        public double NetIncome { get; set; }

        public double NetIncomeRatio { get; set; }

        public double EPS { get; set; }
        public double EPSDiluted { get; set; }

        public double weightedAverageShsOut { get; set; }

        public double weightedAverageShsOutDil { get; set; }

        public string Link { get; set; }

        public string FinalLink { get; set; }


        //balance sheet

        public string FilingDate { get; set; }

        public string AcceptedDate { get; set; }

        public string CalendarYear { get; set; }

        public string Period { get; set; }

        public double CashAndCashEquivalents { get; set; }

        public double ShortTermInvestments { get; set; }

        public double NetReceivables { get; set; }

        public double Inventory { get; set; }
        public double OtherCurrentAssets { get; set; }
        public double TotalCurrentAssets { get; set; }
        public double propertyPlantEquipmentNet { get; set; }
        public double Goodwill { get; set; }
        public double IntangibleAssets { get; set; }
        public double GoodwillAndIntangibleAssets { get; set; }
        public double LongtermInvestments { get; set; }
        public double TaxAssets { get; set; }
        public double OtherNonCurrentAssets { get; set; }
        public double TotalNonCurrentassets { get; set; }
        public double OtherAssets { get; set; }
        public double TotalAssets { get; set; }
        public double AccountPayables { get; set; }
        public double ShortTermDebt { get; set; }
        public double TaxPayables { get; set; }
        public double DeferredRevenue { get; set; }

        public double OtherCurrentLiabilities { get; set; }
        public double totalCurrentLiabilities { get; set; }
        public double LongTermDebt { get; set; }
        public double DeferredRevenueNonCurrent { get; set; }
        public double DeferredTaxLiabilitiesNonCurrent { get; set; }
        public double OtherNonCurrentLiabilities { get; set; }
        public double TotalNonCurrentLiabilities { get; set; }
        public double OtherLiabilities { get; set; }
        public double CapitalLeaseObligations { get; set; }
        public double totalLiabilities { get; set; }
        public double PreferredStock { get; set; }
        public double CommonStock { get; set; }
        public double RetainedEarnings { get; set; }
        public double AccumulatedOtherComprehensiveIncomeLoss { get; set; }
        public double OtherTotalStockholdersEquity { get; set; }
        public double TotalStockholdersEquity { get; set; }
        public double TotalLiabilitiesAndStockholdersEquity { get; set; }
        public double MinorityInterest { get; set; }
        public double TotalEquity { get; set; }
        public double TotalLiabilitiesAndTotalEquity { get; set; }
        public double TotalInvestments { get; set; }
        public double TotalDebt { get; set; }

        public double NetDebt { get; set; }

        public double DeferredIncomeTax { get; set; }

        public double StockBasedCompensation { get; set; }

        public double ChangeInWorkingCapital { get; set; }

        public double AccountsReceivables { get; set; }
        public double AccountsPayables { get; set; }
        public double OtherWorkingCapital { get; set; }

        public double OtherNonCashItems { get; set; }
        public double NetCashProvidedByOperatingActivities { get; set; }
        public double InvestmentsInPropertyPlantAndEquipment { get; set; }
        public double AcquisitionsNet { get; set; }
        public double PurchasesOfInvestments { get; set; }
        public double SalesMaturitiesOfInvestments { get; set; }
        public double otherInvestingActivites { get; set; }
        public double netCashUsedForInvestingActivites { get; set; }
        public double DebtRepayment { get; set; }
        public double CommonStockIssued { get; set; }
        public double CommonStockRepurchased { get; set; }
        public double DividendsPaid { get; set; }
        public double otherFinancingActivites { get; set; }
        public double NetCashUsedProvidedByFinancingActivities { get; set; }
        public double EffectOfForexChangesOnCash { get; set; }
        public double NetChangeInCash { get; set; }
        public double CashAtEndOfPeriod { get; set; }
        public double cashAtBeginningOfPeriod { get; set; }
        public double OperatingCashFlow { get; set; }
        public double CapitalExpenditure { get; set; }
        public double FreeCashFlow { get; set; }
    }
}
