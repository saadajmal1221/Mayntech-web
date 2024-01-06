namespace Mayntech___Individual_Solution.Auxiliar.Auxiliar.FMP
{
    public class StockSymbolsList
    {
        public List<string> symbols { get; set; }

    }
    public class CompaniesStockListFinal
    {
        public string symbol { get; set; }
        public string name { get; set; }
    }
    public class CompanyProfileFinal
    {
        public string symbol { get; set; }

        public string companyName { get; set; }
        public string currency { get; set; }
        public string isin { get; set; }

        public string industry { get; set; }

        public string sector { get; set; }
        public bool isAdr { get; set; }
        public bool isFund { get; set; }

        public bool isEtf { get; set; }

        public bool isActivelyTrading { get; set; }
    }


}
