using Mayntech___Individual_Solution.Auxiliar.Auxiliar.FMP;
using Mayntech___Individual_Solution.Auxiliar_Analysis.Auxiliar.Other;
using Mayntech___Individual_Solution.Auxiliar_Analysis.ValueTraps;
using Mayntech___Individual_Solution.Auxiliar_Valuation.Financials_projections;
using Newtonsoft.Json;
using System.Text.Json;

namespace Mayntech___Individual_Solution.Services
{
    public class JsonService
    {
        public JsonService(IWebHostEnvironment webHostEnvironment)
        {
            WebHostEnvironment = webHostEnvironment;
        }

        public IWebHostEnvironment WebHostEnvironment { get; }

        private string jsonFileName
        {
            get { return Path.Combine(WebHostEnvironment.WebRootPath, "Data", "Taxes.json"); }
        }

        public List<Taxes> GetTaxes()
        {
            using (var jsonFileReader = File.OpenRead(jsonFileName))
            {
                bool fileExists = File.Exists(jsonFileName);
                string json = File.ReadAllText(jsonFileName);
                List<Taxes> taxes = JsonConvert.DeserializeObject<List<Taxes>>(json);

                return taxes;
            }
        }

        private string jsonFileNameCompanyindustry
        {
            get { return Path.Combine(WebHostEnvironment.WebRootPath, "Data", "CompanyIndustry.json"); }
        }

        public List<IndustryDamodaran> GetCompanyindustry()
        {
            using (var jsonFileReader = File.OpenRead(jsonFileNameCompanyindustry))
            {
                bool fileExists = File.Exists(jsonFileNameCompanyindustry);
                string json = File.ReadAllText(jsonFileNameCompanyindustry);
                List<IndustryDamodaran> industry = JsonConvert.DeserializeObject<List<IndustryDamodaran>>(json);

                return industry;
            }
        }


        private string jsonFileWacc
        {
            get { return Path.Combine(WebHostEnvironment.WebRootPath, "Data", "WaccDamodaran.json"); }
        }

        public List<WaccDamodaran> GetCompanyWacc()
        {
            using (var jsonFileReader = File.OpenRead(jsonFileWacc))
            {
                bool fileExists = File.Exists(jsonFileWacc);
                string json = File.ReadAllText(jsonFileWacc);
                List<WaccDamodaran> wacc = JsonConvert.DeserializeObject<List<WaccDamodaran>>(json);

                return wacc;
            }
        }

        private string jsonFileMRP
        {
            get { return Path.Combine(WebHostEnvironment.WebRootPath, "Data", "MarketRiskPremium.json"); }
        }

        public List<MarketRiskP> GetCompanyMRP()
        {
            using (var jsonFileReader = File.OpenRead(jsonFileMRP))
            {
                bool fileExists = File.Exists(jsonFileMRP);
                string json = File.ReadAllText(jsonFileMRP);
                List<MarketRiskP> Mrp = JsonConvert.DeserializeObject<List<MarketRiskP>>(json);

                return Mrp;
            }
        }

        private string jsonFileRoic
        {
            get { return Path.Combine(WebHostEnvironment.WebRootPath, "Data", "RoicRegion.json"); }
        }

        public List<ROIC> GetCompanyROIC()
        {
            using (var jsonFileReader = File.OpenRead(jsonFileRoic))
            {
                bool fileExists = File.Exists(jsonFileRoic);
                string json = File.ReadAllText(jsonFileRoic);
                List<ROIC> roic = JsonConvert.DeserializeObject<List<ROIC>>(json);

                return roic;
            }
        }
    }
}
