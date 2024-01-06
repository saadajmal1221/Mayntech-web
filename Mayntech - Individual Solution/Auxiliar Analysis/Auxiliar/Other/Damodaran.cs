using Mayntech___Individual_Solution.Auxiliar_Analysis.ValueTraps;

namespace Mayntech___Individual_Solution.Auxiliar_Analysis.Auxiliar.Other
{
    public class Damodaran
    {
        public IndustryDamodaran GetCompanyIndustry(List<IndustryDamodaran> industry, string companyName, string companyTick,
            string companyCountry, string companyExchange)
        {
            IndustryDamodaran output = new IndustryDamodaran();



            List<IndustryDamodaran> possibleCompanies = industry.Where(p => p.Ticker == companyTick).ToList();

            if (possibleCompanies.Count()>1)
            {
                List<IndustryDamodaran> firstCheck = possibleCompanies.Where(p => p.CountryIso_2 == companyCountry && p.Exchange == companyExchange).ToList();

                if (firstCheck.Count() > 1)
                {
                    List<IndustryDamodaran> SecondCheck = possibleCompanies.Where(p => p.CompanyName == companyName || p.CompanyName == companyName + " ").ToList();

                    if (SecondCheck.Count() ==1)
                    {
                        output = SecondCheck[0];
                    }
                }
                else if (firstCheck.Count() == 1)
                {
                    output = firstCheck[0];
                }

            }
            else if (possibleCompanies.Count() == 1)
            {
                List<IndustryDamodaran> firstCheck = possibleCompanies.Where(p => p.CountryIso_2 == companyCountry).ToList();

                if (firstCheck.Count()==1)
                {
                    output = firstCheck[0];
                }
            }

            if (output.CompanyName == null)
            {
                List<IndustryDamodaran> possibleCompanies1 = industry.Where(p => p.CompanyName == companyName || p.CompanyName == companyName + " ").ToList();

                if (possibleCompanies1.Count() >1)
                {
                    List<IndustryDamodaran> SecondCheck = possibleCompanies.Where(p => p.CountryIso_2 == companyCountry && p.Exchange == companyExchange).ToList();

                    if (SecondCheck.Count() == 1)
                    {
                        output = SecondCheck[0];
                    }
                }
                else if (possibleCompanies1.Count() == 1)
                {
                    output = possibleCompanies1[0];
                }
                else
                {
                    string companyTickAux = null;
                    try
                    {
                        companyTickAux = companyTick.Substring(0, companyTick.IndexOf("."));
                    }
                    catch (Exception)
                    {

                        companyTickAux = companyTick;
                    }
   
                    

                    List<IndustryDamodaran> firstCheck = possibleCompanies.Where(p => p.Ticker == companyTickAux).ToList();

                    if (firstCheck.Count() == 1)
                    {
                        List<IndustryDamodaran> secondCheck = firstCheck.Where(p => p.CountryIso_2 == companyCountry).ToList();

                        if (firstCheck.Count() == 1)
                        {
                            output = firstCheck[0];
                        }
                    }
                    else if (firstCheck.Count() > 1)
                    {
                        List<IndustryDamodaran> secondCheck = firstCheck.Where(p => p.CountryIso_2 == companyCountry).ToList();

                        if (secondCheck.Count()>1)
                        {
                            List<IndustryDamodaran> thirdCheck = secondCheck.Where(p => p.Exchange == companyExchange).ToList();

                            if (thirdCheck.Count() == 1)
                            {
                                output = thirdCheck[0];
                            }
                            else
                            {

                            }
                        }
                    }
                }
            }

            return output;
        }
    }
    public class WaccDamodaran
    {
        public string Region { get; set; }
        public string IndustryName { get; set; }
        public string NumberofFirms { get; set; }
        public string Beta { get; set; }
        public string CostofEquity { get; set; }
        public string equityToCapital { get; set; }
        public string StdDevinStock { get; set; }
        public string CostofDebt { get; set; }
        public string TaxRate { get; set; }
        public string afterTaxCostofDebt { get; set; }
        public string debtToCapital { get; set; }
        public string CostofCapital { get; set; }
    }
    public class MarketRiskP
    {
        public string Region { get; set; }
        public string MarketRiskPremium { get; set; }
    }
    public class ROIC
    {
        public string Region { get; set; }
        public string IndustryName { get; set; }
        public string ROICwithoutleases { get; set; }
        public string ROICwithleases { get; set; }
    }
}
