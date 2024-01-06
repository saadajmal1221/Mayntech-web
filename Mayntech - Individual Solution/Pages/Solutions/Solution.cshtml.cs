 using Mayntech___Individual_Solution.Auxiliar.Analysis;
using Mayntech___Individual_Solution.Auxiliar.Auxiliar;
using Mayntech___Individual_Solution.Auxiliar.Auxiliar.FMP;
using Mayntech___Individual_Solution.Auxiliar.CompanyAnalysis;
using Mayntech___Individual_Solution.Auxiliar.CompanyAnalysis.Competitors;
using Mayntech___Individual_Solution.Auxiliar.Executive_Summary;
using Mayntech___Individual_Solution.Auxiliar.ValueTraps;
using Mayntech___Individual_Solution.Auxiliar_Analysis.Auxiliar.Other;
using Mayntech___Individual_Solution.Auxiliar_Analysis.CompetitorsRatios;
using Mayntech___Individual_Solution.Auxiliar_Analysis.Valuation;
using Mayntech___Individual_Solution.Auxiliar_Analysis.ValueTraps;
using Mayntech___Individual_Solution.Auxiliar_Valuation.Financials_projections;
using Mayntech___Individual_Solution.Services;
using Microsoft.AspNetCore.Authorization;
using Microsoft.AspNetCore.Mvc;
using Microsoft.AspNetCore.Mvc.RazorPages;
using Microsoft.AspNetCore.Mvc.Rendering;
using Newtonsoft.Json;
using OfficeOpenXml;
using System.ComponentModel.DataAnnotations;
using System.Data.SqlClient;
using System.Text.Json.Serialization;
using Microsoft.Extensions.Configuration;

namespace Mayntech___Individual_Solution.Pages
{

    //[Authorize]
    public class SolutionModel : PageModel
    {

        

        [BindProperty]
        public Ticker tickerInput { get; set; }
        [BindProperty]
        public NumberYears years { get; set; }
        public static int NumberYears = 0;

        public static List<FinancialStatements> incomeStatement = new List<FinancialStatements>();
        public List<FinancialStatements> incomePeers;
        public List<FinancialStatements> balancePeers;
        public List<FinancialStatements> cashPeers;
        public List<FinancialStatements> incomeStatementAsReported;
        public static List<FinancialStatements> balances = new List<FinancialStatements>();
        public static List<FinancialStatements> cashFlow  = new List<FinancialStatements>();
        public CompanyProfile companyOutlook;
        public PeersProfile companyProfile;
        public PeersProfile PeersOutlook;
        public List<CompaniesStockListFinal> companyinfo = new List<CompaniesStockListFinal>();


        public List<PeersList> companypeers;
        public String ErrorMsg = "";

        public SelectCompanyModel selectCompany;
        FinancialMP fmp = new FinancialMP();
        public List<CompaniesStockListFinal> StockList = new List<CompaniesStockListFinal>();
        static public string CompanyTick;

        public List<Taxes> taxes { get; private set; }

        public static IDictionary<string, List<FinancialStatements>> IncomeStatementDict = new Dictionary<string, List<FinancialStatements>>();
        public static IDictionary<string, List<FinancialStatements>> BalanceSheetDict = new Dictionary<string, List<FinancialStatements>>();
        public static IDictionary<string, List<FinancialStatements>> cashFlowDict = new Dictionary<string, List<FinancialStatements>>();



        private readonly ILogger<PageModel> _logger;
        public JsonService jsonService;

        public List<IndustryDamodaran> industry { get; private set; }

        private readonly IConfiguration _config;

        public SolutionModel(ILogger<PageModel> logger, JsonService jsonService, IConfiguration config)
        {
            _logger = logger;
            this.jsonService = jsonService;
            _config = config;
        }


        [HttpGet]
        public IActionResult OnGetComplete(string prefix)
        {
            var environment = _config.GetValue<string>("ASPNETCORE_ENVIRONMENT");

            string connectionString = null;


            connectionString = _config["connectionString"];


            incomeStatement = null;

            List<string> stocklist = new List<string>();
            // Retrieve the connection string
            
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

        public async Task OnGetAsync()
        {
            //var connectionString = _config["ConnectionStrings:AppConfig"];
            //companyinfo = await fmp.GetCompanyInfoAPIAsync(companyinfo, connectionString);

            //DatabaseUpdate update = new DatabaseUpdate();
            //update.Database();

            Screener screener = new Screener();
            //screener.StockScreener();

            //screener.NumberOfSectors();
            //screener.NumberOfIndustries();

            

        }

        //private static void DoWork(CancellationToken token)
        //{
        //    while (true)
        //    {
        //        if (token.IsCancellationRequested)
        //        {
        //            // exit the loop and the function
        //            incomeStatement.Clear();
        //            balances.Clear();
        //            cashFlow.Clear();
        //            IncomeStatementDict.Clear();
        //            BalanceSheetDict.Clear();
        //            cashFlowDict.Clear();
        //            ExecSummary.ExecSummaryDetails.Clear();
        //            break;
        //        }

        //        // do some work here

        //        // check again after a short delay
        //        Thread.Sleep(500);
        //    }
        //}

        public async Task<FileStreamResult> OnPostAsync()
        {
            string companyName = Request.Form["SearchCompany"];

            years.Years = 10;

            var environment = _config.GetValue<string>("ASPNETCORE_ENVIRONMENT");

            string connectionString = null;

            connectionString = _config["connectionString"];

            //CancellationTokenSource cts = new CancellationTokenSource();

            //Task.Run(() => DoWork(cts.Token));

            //cts.Cancel();

            taxes = jsonService.GetTaxes();


            if (incomeStatement !=null)
            {
                incomeStatement.Clear();
            }
            if (balances !=null)
            {
                balances.Clear();
            }
            if (cashFlow != null)
            {
                cashFlow.Clear();
            }
            
            

            try
            {
                String connection = connectionString;
                using (SqlConnection con = new SqlConnection(connection))
                {
                    con.Open();
                    String sql = "SELECT symbol FROM ListOfStocks WHERE name ='" + companyName.Replace("'","").Trim() + "'";
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



            double[] revenue;
            int i = 0;

            //Vai buscar as DFs ao FMP
            try
            {
                string APIIncomeStatement = "https://financialmodelingprep.com/api/v3/income-statement/" + CompanyTick + "?apikey=d7ba591715e23c3d2ac3b0aec6dea138";
                incomeStatement = await fmp.GetFinancialStatementsAPIAsync(APIIncomeStatement);
                
            }
            catch (Exception)
            {

                ErrorMsg = "We do not have financial data regarding this company. Please select another company";
                return null;
            }


            if (incomeStatement.Count() == 0 || incomeStatement[0].Date.Year<DateTime.Now.Year - 2)
            {
                ErrorMsg = "We do not have financial data regarding this company. Please select another company";
                return null;
            }
            else
            {
                ErrorMsg = "";
            }

            DateTime referenceDate = incomeStatement[0].Date;

            string APIPeers = "https://financialmodelingprep.com/api/v4/stock_peers?symbol=" + CompanyTick + "&apikey=d7ba591715e23c3d2ac3b0aec6dea138";
            companypeers = await fmp.GetCompanyPeers(APIPeers);

            string APIBalanceSheet = "https://financialmodelingprep.com/api/v3/balance-sheet-statement/" + CompanyTick + "?apikey=d7ba591715e23c3d2ac3b0aec6dea138";
            balances = await fmp.GetFinancialStatementsAPIAsync(APIBalanceSheet);
            string CalendarYear = balances[0].CalendarYear;

            string APICashFlow = "https://financialmodelingprep.com/api/v3/cash-flow-statement/" + CompanyTick + "?apikey=d7ba591715e23c3d2ac3b0aec6dea138";
            cashFlow = await fmp.GetFinancialStatementsAPIAsync(APICashFlow);

            string APICompanyOutlook = "https://financialmodelingprep.com/api/v4/company-outlook?symbol=" + CompanyTick + "&apikey=d7ba591715e23c3d2ac3b0aec6dea138";
            companyProfile = await fmp.GetPeersOutlook(APICompanyOutlook);

            companyOutlook = await fmp.GetCompanyOutlook(APICompanyOutlook);

            industry = jsonService.GetCompanyindustry();

            List<ROIC> ListRoic = jsonService.GetCompanyROIC();

            Damodaran damodaran = new Damodaran();
            IndustryDamodaran companyIndustry = damodaran.GetCompanyIndustry(industry, companyName, CompanyTick, companyProfile.profile.country, companyProfile.profile.exchangeShortName);
            List<ROIC> roics = new List<ROIC>();

            if (companyIndustry != null)
            {
                try
                {
                    roics = ListRoic.Where(p => p.Region == companyIndustry.BroadGroup && p.IndustryName == companyIndustry.IndustryGroup).ToList();
                }
                catch (Exception)
                {

                    roics = null;
                }
                

            }


            if (companyProfile.profile.isEtf == true || companyProfile.profile.isin == null ||
                companyProfile.profile.sector == "Financial Services" || companyProfile.profile.isActivelyTrading == false || companyProfile.financialsQuarter.income.Count()==0
                || companyProfile.financialsQuarter.balance.Count()==0)
            {
                ErrorMsg = "We do not have financial data regarding this company. Please select another company";
                return null;
            }

            //Vai buscar a lista de empresas dos peers
            CompetitorsList competitorsList = new CompetitorsList();

            List<string> competitors = competitorsList.GetCompetitorsList(companypeers);
            int numberOfCompetitors = competitors.Count();

            int numberOfYearsStatements = Math.Min(Math.Min(incomeStatement.Count(), balances.Count()), cashFlow.Count());

            int numberOfYears = Math.Min(years.Years, numberOfYearsStatements);


            IDictionary<string, List<FinancialStatements>> IncomeStatementDictAUx = new Dictionary<string, List<FinancialStatements>>();
            IDictionary<string, List<FinancialStatements>> BalanceSheetDictAux = new Dictionary<string, List<FinancialStatements>>();
            IDictionary<string, List<FinancialStatements>> cashFlowDictAux = new Dictionary<string, List<FinancialStatements>>();

            Dictionary<string, string> CompetitorsDict = new Dictionary<string, string>();

            for (int a = 0; a < numberOfCompetitors; a++)
            {

                try
                {
                    string APIPeersOutlook = "https://financialmodelingprep.com/api/v4/company-outlook?symbol=" + competitors[a] + "&apikey=d7ba591715e23c3d2ac3b0aec6dea138";
                    PeersOutlook = await fmp.GetPeersOutlook(APIPeersOutlook);

                    CompetitorsDict.Add(PeersOutlook.profile.companyName, PeersOutlook.profile.symbol);

                    NormalizePeersStatements normalizePeersStatements = new NormalizePeersStatements();
                    List<FinancialStatements> peerNormalized = normalizePeersStatements.FinancialStatementsNormalizations(PeersOutlook, referenceDate, numberOfYears);

                    incomePeers = peerNormalized;

                    balancePeers = peerNormalized;

                    cashPeers = peerNormalized;

                    //string APIincomestatementPeers = "https://financialmodelingprep.com/api/v3/income-statement/" + competitors[a] + "?apikey=d7ba591715e23c3d2ac3b0aec6dea138";
                    //incomePeers = await fmp.GetPeersFinancialsAPIAsync(APIincomestatementPeers);

                    //string APIBalanceSheetPeers = "https://financialmodelingprep.com/api/v3/balance-sheet-statement/" + competitors[a] + "?apikey=d7ba591715e23c3d2ac3b0aec6dea138";
                    //balancePeers = await fmp.GetPeersFinancialsAPIAsync(APIBalanceSheetPeers);

                    //string APICashFlowPeers = "https://financialmodelingprep.com/api/v3/cash-flow-statement/" + competitors[a] + "?apikey=d7ba591715e23c3d2ac3b0aec6dea138";
                    //cashPeers = await fmp.GetPeersFinancialsAPIAsync(APICashFlowPeers);

                    IncomeStatementDictAUx.Add(competitors[a], incomePeers);
                    BalanceSheetDictAux.Add(competitors[a], balancePeers);
                    cashFlowDictAux.Add(competitors[a], cashPeers);
                }
                catch (Exception)
                {

                    
                }


            }
            IncomeStatementDict = IncomeStatementDictAUx;
            BalanceSheetDict = BalanceSheetDictAux;
            cashFlowDict = cashFlowDictAux;

            Taxes? tax = taxes.FirstOrDefault(t => t.iso_2 == companyProfile.profile.country);

            //Cria o ficheiro de excel e exporta para o user

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

                NumberYears = numberOfYears;
                int numberOfYearsIncomeStatement = incomeStatement.Count();


                Disclaimer disclaimer = new Disclaimer();
                await disclaimer.CreateDisclaimer(package);

                SupportSheet support = new SupportSheet();
                support.supportConstr(package, CompetitorsDict);

                CreateSeparators separators = new CreateSeparators();
                await separators.CreateSeparator(package, "Company Overview ->");

                CompanyOverview overview = new CompanyOverview();
                overview.CompanyOverviewConstruction(companyOutlook, package);
                IncomeStatementConstruction IS = new IncomeStatementConstruction();
                await IS.CreatePL(package, incomeStatement, col, row, companyName, companyProfile);
                BSConstruction bSConstruction = new BSConstruction();
                await bSConstruction.CreateBS(package, balances, col, row, companyName, companyProfile);
                CashFlowStatementConstruction CFConstruction = new CashFlowStatementConstruction();
                await CFConstruction.CreateBS(package, cashFlow, col, row, companyName);

                totalNumberPeriodsCFS = cashFlow.Count();
                totalNumberPeriods = balances.Count();
                totalNumberPeriodsIS = incomeStatement.Count();

                columnsToDeleteCFS = totalNumberPeriodsCFS - numberOfYears;
                for (int a = 1; a <= columnsToDeleteCFS; a++)
                {
                    package.Workbook.Worksheets["CFS"].DeleteColumn(3);
                }

                columnsToDelete = totalNumberPeriods - numberOfYears;
                for (int b = 1; b <= columnsToDelete; b++)
                {
                    package.Workbook.Worksheets["BS"].DeleteColumn(3);
                }

                columnsToDelete = totalNumberPeriodsIS - numberOfYears;
                for (int b = 1; b <= columnsToDelete; b++)
                {
                    package.Workbook.Worksheets["P&L"].DeleteColumn(3);
                }

                await separators.CreateSeparator(package, "Company Analysis ->");

                CompanyAnalysis analysis = new CompanyAnalysis();
                analysis.financialSummary(package, numberOfYears, incomeStatement, balances, cashFlow, tax);

                //WorkingCapital workingCapital = new WorkingCapital();
                //workingCapital.WorkingCapitalConstruction(package, numberOfYears, companyName);

                try
                {
                    CompetitorsAnalysis competitorsAnalysis = new CompetitorsAnalysis();
                    competitorsAnalysis.IncomeStatementCompetitors(package, companyName, numberOfYears);
                    //competitorsAnalysis.BalanceSheetCompetitors(package, companyName, numberOfYears);
                }
                catch 
                {


                }



                await separators.CreateSeparator(package, "Value Analysis ->");

                AuxiliarAnalysis auxPage = new AuxiliarAnalysis();
                auxPage.AuxiliarConstruction(package, numberOfYears);

                //ValueIndex valueIndex = new ValueIndex();
                //valueIndex.ValueTrapsConstruction(package, companyName);

                //ROICAndGrowth roic = new ROICAndGrowth();
                //roic.roicContruction(package, numberOfYears, IncomeStatementDict, BalanceSheetDict, companyName, CalendarYear);

                Growth growth = new Growth();
                growth.GrowthBuilder(package, numberOfYears, "Analysis", roics);

                await separators.CreateSeparator(package, "Ratios ->");

                RatiosAnalysis competitorsanalysis = new RatiosAnalysis();
                competitorsanalysis.RatioConstruction(package, numberOfYears, IncomeStatementDict, BalanceSheetDict, cashFlowDict, numberOfCompetitors, incomeStatement, balances, cashFlow, CompanyTick, numberOfYearsIncomeStatement);

                await separators.CreateSeparator(package, "Valuation Support ->");

                List<FinancialStatements> BalanceSheet = (List<FinancialStatements>)balances.GetRange(0, numberOfYears);
                ValuationSupport valuationSupport = new ValuationSupport();
                valuationSupport.ValuationSupportConstruction(package, numberOfYears, tax, CalendarYear, BalanceSheet);

                //AnalysisIS analysisPL = new AnalysisIS();
                //await analysisPL.CreateISAnalysis(package, incomeStatement, col, row, companyName, numberOfYears);

                //AnalysisBS analysisBS = new AnalysisBS();
                //await analysisBS.CreateBSAnalysis(package, balances, incomeStatement, col, row, companyName, numberOfYears);

                ExecSummaryConstruction execSummary = new ExecSummaryConstruction();
                execSummary.ConstructionExecSummary(package, companyName);

                package.Workbook.Worksheets.MoveAfter("Executive Summary", "Disclaimer");
                BalanceSheetDict.Clear();
                cashFlowDict.Clear();
                IncomeStatementDict.Clear();


                //Gravar o excel final

                package.Save();
            }

            //Exportar o excel

            stream.Position = 0;
            string excelName = "MaynTech Analysis - "+ companyName + ".xlsx";
            return File(stream, "application/octet-stream", excelName);
        }
        
    }
    public class Ticker
    {
        [Required]
        [Display(Name = "Ticker")]
        public string tick { get; set; }
    }

    public class EmpresasModel
    {

        [JsonPropertyName("Empresas")]
        public List<string>? Empresas { get; set; }
    }
    public class SelectCompanyModel
    {
        public string SelectedCompany { get; set; }
        //display property
        public List<SelectListItem> SelectListaEmpresas { get; set; }
    }
    public class NumberYears
    {
        [Required]
        [Display(Name = "Number of years")]
        public int Years { get; set; }
    }


}
