using Mayntech___Individual_Solution.Auxiliar.Analysis;
using Mayntech___Individual_Solution.Auxiliar.Auxiliar;
using Mayntech___Individual_Solution.Auxiliar.Auxiliar.FMP;
using Mayntech___Individual_Solution.Auxiliar.CompanyAnalysis.Competitors;
using Mayntech___Individual_Solution.Auxiliar.CompanyAnalysis.WorkingCapital;
using Mayntech___Individual_Solution.Auxiliar.CompanyAnalysis;
using Mayntech___Individual_Solution.Auxiliar_Analysis.CompetitorsRatios;
using Mayntech___Individual_Solution.Auxiliar.Executive_Summary;
using Mayntech___Individual_Solution.Auxiliar.ValueTraps;
using Microsoft.AspNetCore.Mvc;
using Microsoft.AspNetCore.Mvc.RazorPages;
using Microsoft.AspNetCore.Mvc.Rendering;
using OfficeOpenXml;
using System.ComponentModel.DataAnnotations;
using System.Data.SqlClient;
using System.IO;
using System.Xml.Linq;
using Mayntech___Individual_Solution.Auxiliar_Analysis.ValueTraps;
using Mayntech___Individual_Solution.Auxiliar_Valuation.Financials_projections;
using Mayntech___Individual_Solution.Services;
using System.Data;
using Microsoft.Extensions.Logging;
using Mayntech___Individual_Solution.Auxiliar_Analysis.Auxiliar.FMP;
using Mayntech___Individual_Solution.Auxiliar_Analysis.Auxiliar.Other;
using Mayntech___Individual_Solution.Auxiliar_Valuation.Assumptions;
using Mayntech___Individual_Solution.Auxiliar_Valuation.Multiples_Valuation;
using Microsoft.JSInterop;
using Azure;
using Microsoft.AspNetCore.Authorization;

namespace Mayntech___Individual_Solution.Pages.Solutions
{
    //[Authorize]
    public class ValuationModel : PageModel
    {

        public List<FinancialStatements> incomeStatementVal = new List<FinancialStatements>();

        public String ErrorMsg = "";

        

        
        public string CompanyTick { get; set; }
        public string companyName { get; set; }

        private readonly IConfiguration _config;
        private readonly ILogger<PageModel> _logger;
        

        public ValuationModel(IConfiguration config, JsonService jsonService, ILogger<PageModel> logger)
        {
            _logger = logger;
            this.jsonService = jsonService;
            _config = config;
        }

        public List<FinancialStatements> balancesVal = new List<FinancialStatements>();
        public List<FinancialStatements> cashFlowVal = new List<FinancialStatements>();
        

        public List<IndustryDamodaran> industry { get; private set; }

        public List<MarketRiskPremium> marketRiskPremia = new List<MarketRiskPremium>();
        public List<CompanyNotes> companyNotes = new List<CompanyNotes>();
        public CompanyProfile companyOutlook;

        public List<PeersList> companypeers;

        public Dictionary<string, List<FinancialStatements>> financialsComparables = new Dictionary<string, List<FinancialStatements>>();

        public PeersProfile PeersOutlook;
        public PeersProfile CompanyProfile;

        public List<PeersProfile> PeersProfileList = new List<PeersProfile>();
        

        public DataTable dataTable = new DataTable();

        public JsonService jsonService;

        public List<Taxes> taxes { get; private set; }


        //public void OnGet()
        //{           
        //    CompanyTick = null;
        //    ErrorMsg = "";
        //}


        [HttpGet]
        public IActionResult OnGetSelect(string prefix)
        {
            

            var environment = _config.GetValue<string>("ASPNETCORE_ENVIRONMENT");

            string connectionString = null;

            connectionString = _config["connectionString"];

            List<string> stocklist = new List<string>();
            String connection = connectionString;
            using (SqlConnection conn = new SqlConnection(connection))
            {
                conn.Open();

                string select = "SELECT * FROM ListOfStocks WHERE name LIKE '%" + prefix + "%'";

                using (SqlCommand cmd = new SqlCommand(select, conn))
                {

                    using (SqlDataReader reader = cmd.ExecuteReader())
                    {
                        while (reader.Read())
                        {

                            stocklist.Add(reader.GetString(1));
                        }

                    }
                    conn.Close();
                }
            }
            return new JsonResult(stocklist);
        }
 
        public async Task<FileStreamResult> OnPostAsync()
        {



            var connectionString = _config["connectionString"];

            companyName = Request.Form["SearchCompany"];

            List<IndustryDamodaran> industry = jsonService.GetCompanyindustry();

            taxes = jsonService.GetTaxes();

            FinancialMP fmp = new FinancialMP();
            try
            {
                String connection = connectionString;
                using (SqlConnection con = new SqlConnection(connection))
                {
                    con.Open();
                    String sql = "SELECT symbol FROM ListOfStocks WHERE name ='" + companyName.Replace("'", "").Trim() + "'";
                    using (SqlCommand command = new SqlCommand(sql, con))
                    {
                        using (SqlDataReader reader = command.ExecuteReader())
                        {
                            reader.Read();
                            CompanyTick = reader.GetString(0);
                        }
                    }

                }
            }
            catch (Exception)
            {

                ErrorMsg = "We do not have financial data regarding this company. Please select another company";
                
                return null;

            }


            string APIBalanceSheet = "https://financialmodelingprep.com/api/v3/balance-sheet-statement/" + CompanyTick + "?apikey=d7ba591715e23c3d2ac3b0aec6dea138";
            balancesVal = await fmp.GetFinancialStatementsAPIAsync(APIBalanceSheet);


            try
            {
                string APIIncomeStatement = "https://financialmodelingprep.com/api/v3/income-statement/" + CompanyTick + "?apikey=d7ba591715e23c3d2ac3b0aec6dea138";
                incomeStatementVal = await fmp.GetFinancialStatementsAPIAsync(APIIncomeStatement);
            }
            catch 
            {
                ErrorMsg = "We do not have financial data regarding this company. Please select another company";
                
                return null;
            }


            if (balancesVal.Count()==0 || incomeStatementVal.Count() == 0 || incomeStatementVal[0].Date.Year < DateTime.Now.Year - 2)
            {
                ErrorMsg = "We do not have financial data regarding this company. Please select another company";
                
                return null;
            }

            string APICashFlow = "https://financialmodelingprep.com/api/v3/cash-flow-statement/" + CompanyTick + "?apikey=d7ba591715e23c3d2ac3b0aec6dea138";
            cashFlowVal = await fmp.GetFinancialStatementsAPIAsync(APICashFlow);

            string APImarketRisk = "https://financialmodelingprep.com/api/v4/market_risk_premium?apikey=d7ba591715e23c3d2ac3b0aec6dea138";
            marketRiskPremia = await fmp.GetMarketRiskPremium(APImarketRisk);

            string APICompanyOutlook = "https://financialmodelingprep.com/api/v4/company-outlook?symbol=" + CompanyTick + "&apikey=d7ba591715e23c3d2ac3b0aec6dea138";
            CompanyProfile = await fmp.GetPeersOutlook(APICompanyOutlook);

            companyOutlook = await fmp.GetCompanyOutlook(APICompanyOutlook);

            string APIcompanyNotes = "https://financialmodelingprep.com/api/v4/company-notes?symbol=" + CompanyTick + "&apikey=d7ba591715e23c3d2ac3b0aec6dea138";
            companyNotes = await fmp.GetCompanyNotes(APIcompanyNotes);

            string APIPeers = "https://financialmodelingprep.com/api/v4/stock_peers?symbol=" + CompanyTick + "&apikey=d7ba591715e23c3d2ac3b0aec6dea138";
            companypeers = await fmp.GetCompanyPeers(APIPeers);


            if (CompanyProfile.profile.isEtf == true || CompanyProfile.profile.isin == null ||
                CompanyProfile.profile.sector == "Financial Services" || CompanyProfile.profile.isActivelyTrading == false || CompanyProfile.financialsQuarter.income.Count() == 0
                || CompanyProfile.financialsQuarter.balance.Count() == 0)
            {
                ErrorMsg = "We do not have financial data regarding this company. Please select another company";
                return null;
            }


            CompetitorsList competitorsList = new CompetitorsList();

            List<string> competitors = competitorsList.GetCompetitorsList(companypeers);
            int numberOfCompetitors = competitors.Count();

            DateTime referenceDate = incomeStatementVal[0].Date;

            int numberOfYears = Math.Min(10, Math.Min(incomeStatementVal.Count(), balancesVal.Count()));

            Dictionary<string, string> competitorsDict = new Dictionary<string, string>();

            for (int a = 0; a < numberOfCompetitors; a++)
            {
                string APIPeersOutlook = "https://financialmodelingprep.com/api/v4/company-outlook?symbol=" + competitors[a] + "&apikey=d7ba591715e23c3d2ac3b0aec6dea138";
                PeersOutlook = await fmp.GetPeersOutlook(APIPeersOutlook);

                try
                {
                    competitorsDict.Add(PeersOutlook.profile.companyName, PeersOutlook.profile.symbol);
                    PeersProfileList.Add(PeersOutlook);
                }
                catch (Exception)
                {

                    
                }
                

                //NormalizePeersStatements normalizePeersStatements = new NormalizePeersStatements();
                //List<FinancialStatements> peerNormalized = normalizePeersStatements.FinancialStatementsNormalizations(PeersOutlook, referenceDate, numberOfYears);

                //financialsComparables.Add(PeersOutlook.profile.symbol, peerNormalized);

                
            }

            var stream = new MemoryStream();
            ExcelPackage.LicenseContext = LicenseContext.NonCommercial;
            using (var package = new ExcelPackage(stream))
            {

                int col = 0;
                int row = 0;

                //Processo para apagar as colunas do excel para os anos que não queremos

                int totalNumberPeriodsCFS;
                int totalNumberPeriods;
                int totalNumberPeriodsIS;
                int columnsToDeleteCFS;
                int columnsToDelete;
                int aux = 4;



                if (numberOfYears < incomeStatementVal.Count())
                {
                    try
                    {
                        incomeStatementVal.RemoveRange(10, incomeStatementVal.Count() - 10);
                        balancesVal.RemoveRange(10, balancesVal.Count() - 10);
                        cashFlowVal.RemoveRange(10, cashFlowVal.Count() - 10);
                    }
                    catch
                    {
                        incomeStatementVal.RemoveRange(numberOfYears, incomeStatementVal.Count() - numberOfYears);

                    }

                }
                else if (incomeStatementVal.Count() > balancesVal.Count())
                {
                    incomeStatementVal.RemoveRange(balancesVal.Count(), incomeStatementVal.Count() - balancesVal.Count());
                }
                else if (balancesVal.Count() > incomeStatementVal.Count())
                {
                    balancesVal.RemoveRange(incomeStatementVal.Count(), balancesVal.Count() - incomeStatementVal.Count());
                }

                int year = incomeStatementVal[0].Date.Year;

                string currency = incomeStatementVal[0].ReportedCurrency;

                Taxes? tax = taxes.FirstOrDefault(t => t.iso_2 == CompanyProfile.profile.country);
                string calendarYear = balancesVal[0].CalendarYear;

                Disclaimer disclaimer = new Disclaimer();
                await disclaimer.CreateDisclaimer(package);

                SupportSheetValuation support = new SupportSheetValuation();
                support.supportConstrVal(package, competitorsDict);

                CreateSeparators separators = new CreateSeparators();
                await separators.CreateSeparator(package, "Reported Financial Statements ->");


                CompanyOverview companyOverview = new CompanyOverview();
                companyOverview.CompanyOverviewConstruction(companyOutlook, package);

                IncomeStatementConstruction IS = new IncomeStatementConstruction();
                await IS.CreatePL(package, incomeStatementVal, col, row, companyName, CompanyProfile);

                int quarters = IS.Quarters(CompanyProfile);

                BSConstruction bSConstruction = new BSConstruction();
                await bSConstruction.CreateBS(package, balancesVal, col, row, companyName, CompanyProfile);
                CashFlowStatementConstruction CFConstruction = new CashFlowStatementConstruction();
                await CFConstruction.CreateBS(package, cashFlowVal, col, row, companyName);

                await separators.CreateSeparator(package, "Auxiliar Sheets->");

                Damodaran damodaran = new Damodaran();
                IndustryDamodaran industryDamodaran = damodaran.GetCompanyIndustry(industry, companyName, incomeStatementVal[0].Symbol, CompanyProfile.profile.country, CompanyProfile.profile.exchangeShortName);
                WaccDamodaran WaccInputsDamodaran = new WaccDamodaran();

                MarketRiskP marketRisk = new MarketRiskP();

                if (industryDamodaran != null)
                {
                    try
                    {
                        string region = industryDamodaran.BroadGroup;
                        List<WaccDamodaran> waccList = jsonService.GetCompanyWacc();

                        List<WaccDamodaran> wacc = waccList.Where(p => p.Region == region && p.IndustryName == industryDamodaran.IndustryGroup).ToList();

                        //marketRisk = mrp.Where(p => p.Region == region).ToList()[0];

                        WaccInputsDamodaran = wacc[0];
                    }
                    catch
                    {

                        WaccInputsDamodaran = null;
                    }

                }


                Assumptions assumptions = new Assumptions();
                await assumptions.assumptionsBuilderAsync(package, incomeStatementVal, numberOfYears, companyOutlook, marketRiskPremia,
                    companyNotes, tax, WaccInputsDamodaran, quarters);

                AuxiliarSheet auxiliarSheet = new AuxiliarSheet();
                auxiliarSheet.AuxiliarSheetConstruction(package, balancesVal, incomeStatementVal, numberOfYears);

                await separators.CreateSeparator(package, "Free Cash Flow Valuation->");

                FCFProjections IstatementProjections = new FCFProjections();
                IstatementProjections.IncomeProjectionConstruction(package, numberOfYears, tax, calendarYear);

                Growth growth = new Growth();
                growth.GrowthBuilder(package, numberOfYears, "Valuation", null);


                ForexList forex = await fmp.GetFxRate();

                Valuation valuation = new Valuation();
                valuation.ValuationBuilder(package, numberOfYears, companyOutlook, currency, forex);

                await separators.CreateSeparator(package, "Multiples Valuation->");

                MultiplesValuation multiples = new MultiplesValuation();
                multiples.MultiplesValuationConstruction(package, year, numberOfYears, PeersProfileList);

                //Gravar o excel final

                package.Save();
            }



            //Exportar o excel

            stream.Position = 0;

            string excelName = "MaynTech Valuation - " + companyName + ".xlsx";




            return File(stream, "application/octet-stream", excelName);
            //return RedirectToPage("/Solutions/Valuation");

        }


    }

    public class valuation
    {
        public string companyName { get; set; }
        public string companyTick { get; set; }
    }
}
