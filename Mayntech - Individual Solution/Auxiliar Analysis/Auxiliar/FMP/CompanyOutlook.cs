namespace Mayntech___Individual_Solution.Auxiliar.Auxiliar.FMP
{
    public class CompanyOutlook
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
    public class CompanyProfile
    {
        public CompanyOutlook profile { get; set; }
    }

    public class MarketRiskPremium
    {
        public string country { get; set; }
        public string continent { get; set; }
        public double totalEquityRiskPremium { get; set; }
        public double countryRiskPremium { get; set; }
    }

    public class TreasuryRates
    {
        public string date { get; set; }

        public decimal year10 { get; set; }
    }

    public class CompanyNotes
    {
        public string cik { get; set; }

        public string symbol { get; set; }
        public string title { get; set; }
        public string exchange { get; set; }
    }

    public class StockScreener
    {
        public string symbol { get; set; }

        public string companyName { get; set; }
        public long marketCap { get; set; }
        public string sector { get; set; }
        public decimal beta { get; set; }
        public double price { get; set; }
        public double lastAnuualDividend { get; set; }
        public long volume { get; set; }
        public string exchange { get; set; }
        public string exchangeShortName { get; set; }
    }
}
