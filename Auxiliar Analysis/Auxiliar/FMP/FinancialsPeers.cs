namespace Mayntech___Individual_Solution.Auxiliar.Auxiliar.FMP
{
    public class FinancialsPeers
    {
        public DateTime Date { get; set; }
        public double ShortTermInvestments { get; set; }
        public double CashAndCashEquivalents { get; set; }
        public double totalCurrentAssets { get; set; }
        public double totalCurrentLiabilities { get; set; }

        public double NetReceivables { get; set; }
        public double Inventory { get; set; }

        public double TotalCurrentAssets { get; set; }
        public double propertyPlantEquipmentNet { get; set; }

        public double Goodwill { get; set; }
        public double IntangibleAssets { get; set; }

        public double LongtermInvestments { get; set; }

        public double AccountPayables { get; set; }
        public double ShortTermDebt { get; set; }

        public double LongTermDebt { get; set; }

        public double TotalDebt { get; set; }
        public double TotalAssets { get; set; }

        public double TotalEquity { get; set; }

        public double OtherCurrentLiabilities { get; set; }
        public double OtherNonCurrentLiabilities { get; set; }

        public double Revenue { get; set; }

        public double GrossProfitRatio { get; set; }


        public double EBITDARatio { get; set; }
        public double OperatingIncomeRatio { get; set; }
        public double NetIncomeRatio { get; set; }
        public double OperatingIncome { get; set; }
        public double OtherCurrentAssets { get; set; }
        public double DeferredRevenue { get; set; }

    }

    public class PeersOutlook
    {
        public string symbol { get; set; }

        public decimal price { get; set; }

        public decimal beta { get; set; }
        public double volAvg { get; set; }
        public double mktCap { get; set; }
        public double lastDiv { get; set; }

        public string companyName { get; set; }
        public string currency { get; set; }
        public string isin { get; set; }
        public string exchangeShortName { get; set; }

        public string industry { get; set; }

        public string description { get; set; }
        public string ceo { get; set; }
        public string sector { get; set; }
        public string country { get; set; }

        public string fullTimeEmployees { get; set; }

        public bool isEtf { get; set; }

        public bool isActivelyTrading { get; set; }




    }

    public class PeersFinancialStatements
    {
        public List<FinancialStatements> income { get; set; }

        public List<FinancialStatements> balance { get; set; }

        public List<FinancialStatements> cash { get; set; }
    }

    public class PeersProfile
    {
        public PeersFinancialStatements financialsAnnual { get; set; }

        public PeersFinancialStatements financialsQuarter { get; set; }

        public PeersOutlook profile { get; set; }
    }
}
