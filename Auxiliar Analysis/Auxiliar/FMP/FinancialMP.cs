using Mayntech___Individual_Solution.Auxiliar_Analysis.Auxiliar.FMP;
using Mayntech___Individual_Solution.Pages;
using Microsoft.AspNetCore.Mvc;
using Microsoft.AspNetCore.Mvc.Rendering;
using Newtonsoft.Json;
using Newtonsoft.Json.Linq;
using System.Data.SqlClient;
using System.Text.Json.Nodes;


namespace Mayntech___Individual_Solution.Auxiliar.Auxiliar.FMP
{
    public class FinancialMP
    {


        public string ErrorMsg;
        public async Task<List<FinancialStatements>> GetFinancialStatementsAPIAsync(string APIkey)
        {
            //Vai buscar a info sobre a empresa ao Financial modeling prep
            using (var httpClient = new HttpClient())

            {
                //Vai buscar as DFs

                using (var request = new HttpRequestMessage(new HttpMethod("GET"), APIkey))

                {
                    request.Headers.TryAddWithoutValidation("Upgrade-Insecure-Requests", "1");

                    var response = await httpClient.SendAsync(request);
                    string jsonResponse = await response.Content.ReadAsStringAsync();

                    if (jsonResponse == null)
                    {

                        return null;
                    }

                    JsonNode forecastNode = JsonNode.Parse(jsonResponse)!;

                    List<FinancialStatements> income = JsonConvert.DeserializeObject<List<FinancialStatements>>(jsonResponse);
                    return income;
                }
            }
        }


        //Finacials Peers
        public async Task<List<FinancialsPeers>> GetPeersFinancialsAPIAsync(string APIkey)
        {
            //Vai buscar a info sobre a empresa ao Financial modeling prep
            using (var httpClient = new HttpClient())

            {
                //Vai buscar as DFs

                using (var request = new HttpRequestMessage(new HttpMethod("GET"), APIkey))

                {
                    request.Headers.TryAddWithoutValidation("Upgrade-Insecure-Requests", "1");

                    var response = await httpClient.SendAsync(request);
                    string jsonResponse = await response.Content.ReadAsStringAsync();

                    if (jsonResponse == null)
                    {

                        return null;
                    }

                    JsonNode forecastNode = JsonNode.Parse(jsonResponse)!;

                    List<FinancialsPeers> income = JsonConvert.DeserializeObject<List<FinancialsPeers>>(jsonResponse);
                    return income;
                }
            }
        }
        //Conecta com a DB - mais tarde isto estará no Azure
        public async Task<List<CompaniesStockListFinal>> GetCompanyInfoAPIAsync( List<CompaniesStockListFinal> companyInfo, string connectionStringInput)
        {

            using (var httpClient = new HttpClient())

            {
                using (var request = new HttpRequestMessage(new HttpMethod("GET"), "https://financialmodelingprep.com/api/v3/financial-statement-symbol-lists?apikey=d7ba591715e23c3d2ac3b0aec6dea138"))

                {
                    request.Headers.TryAddWithoutValidation("Upgrade-Insecure-Requests", "1");

                    var response = await httpClient.SendAsync(request);
                    string jsonResponse = await response.Content.ReadAsStringAsync();

                    JsonNode forecastNode = JsonNode.Parse(jsonResponse)!;

                    List<string> info = JsonConvert.DeserializeObject<List<string>>(jsonResponse);

                    string connectionString = connectionStringInput;

                    using (SqlConnection connection = new SqlConnection(connectionString))
                    {

                        connection.Open();


                        //Cria uma nova companyInfo e ListOfStocks
                        string sqlcompanyInfo = "INSERT INTO companyInfo" +
                            "(symbol, name, price, exchange, type) VALUES" +
                            "(@symbol, @name, @price, @exchange, @type);";

                        string sqlListOfStocks = "INSERT INTO ListOfStocks" +
                            "(symbol, name) VALUES" + "(@symbol, @name);";

                        //string sqlcompanyOutlook = "INSERT INTO companyOutlook" +
                        //    "(symbol, name, currency, isin , industry," +
                        //    " sector, country, isEtf, " +
                        //    "isActivelyTrading) VALUES" +
                        //    "(@symbol,  @name, @currency, @isin , @industry" +
                        //    ", @sector, @country, @isEtf," +
                        //    "@isActivelyTrading);";

                        foreach (var item in info)
                        {

                            //using (SqlCommand command = new SqlCommand(sqlListOfStocks, connection))
                            //{
                            //    try
                            //    {
                            //        command.Parameters.AddWithValue("@symbol", item.symbol);
                            //        command.Parameters.AddWithValue("@name", item.name.Replace("'", "").Trim());

                            //        command.ExecuteNonQuery();
                            //    }
                            //    catch
                            //    {
                            //        continue;
                            //    }

                            //}



                            //Este é a maneira certa de fazer as coisas. Temos de refazer a base dados ligado à internet


                            string APICompanyProfile = "https://financialmodelingprep.com/api/v3/profile/" + item + "?apikey=d7ba591715e23c3d2ac3b0aec6dea138";
                            List<CompanyProfileFinal> companyProfile = new List<CompanyProfileFinal>();
                            companyProfile = await GetCompanyProfile(APICompanyProfile);

                            try
                            {
                                if (companyProfile[0].sector != "Financial Services" && companyProfile[0].isActivelyTrading == true
                                    && companyProfile[0].isin != null && companyProfile[0].isFund == false && companyProfile[0].isEtf == false
                                    && companyProfile[0].isAdr == false)
                                {

                                    using (SqlCommand command = new SqlCommand(sqlListOfStocks, connection))
                                    {
                                        try
                                        {
                                            command.Parameters.AddWithValue("@symbol", item);
                                            command.Parameters.AddWithValue("@name", companyProfile[0].companyName.Replace("'", "").Trim());

                                            command.ExecuteNonQuery();
                                        }
                                        catch
                                        {
                                            continue;
                                        }


                                    }
                                }
                            }
                            catch (Exception)
                            {


                            }


                        }
                    }
                    

                    return companyInfo;
                }
            }
        }
        public async Task<SelectCompanyModel> GetCompanyListAPIAsync(List<string> stocklist)
        {

            using (var httpClient = new HttpClient())

            {
                var obj = new EmpresasModel();
                using (var request = new HttpRequestMessage(new HttpMethod("GET"), "https://financialmodelingprep.com/api/v3/available-traded/list?apikey=d7ba591715e23c3d2ac3b0aec6dea138"))

                {
                    request.Headers.TryAddWithoutValidation("Upgrade-Insecure-Requests", "1");

                    var response = await httpClient.SendAsync(request);
                    string jsonResponse = await response.Content.ReadAsStringAsync();

                    obj.Empresas = JsonConvert.DeserializeObject<List<string>>(jsonResponse);
                    obj.Empresas.Sort();

                    var model = new SelectCompanyModel();
                    model.SelectListaEmpresas = new List<SelectListItem>();

                    foreach (var item in obj.Empresas)
                    {
                        if (stocklist.Contains(item))
                        {
                            if (true)
                            {
                                model.SelectListaEmpresas.Add(new SelectListItem { Text = item });
                            }

                        }
                    }

                    return model;

                }
            }
        }
        public async Task<CompanyProfile> GetCompanyOutlook(string APIkey)
        {
            //Vai buscar a info sobre a empresa ao Financial modeling prep
            using (var httpClient = new HttpClient())

            {
                //Vai buscar as DFs

                using (var request = new HttpRequestMessage(new HttpMethod("GET"), APIkey))

                {
                    request.Headers.TryAddWithoutValidation("Upgrade-Insecure-Requests", "1");

                    var response = await httpClient.SendAsync(request);
                    string jsonResponse = await response.Content.ReadAsStringAsync();

                    if (jsonResponse == null)
                    {

                        return null;
                    }

                    JsonNode forecastNode = JsonNode.Parse(jsonResponse)!;

                    CompanyProfile companyProfilesOutput = JsonConvert.DeserializeObject<CompanyProfile>(jsonResponse);


                    return companyProfilesOutput;

                }
            }
        }


        public async Task<List<PeersList>> GetCompanyPeers(string APIkey)
        {
            //Vai buscar a info sobre a empresa ao Financial modeling prep
            using (var httpClient = new HttpClient())

            {
                //Vai buscar as DFs

                using (var request = new HttpRequestMessage(new HttpMethod("GET"), APIkey))

                {
                    request.Headers.TryAddWithoutValidation("Upgrade-Insecure-Requests", "1");

                    var response = await httpClient.SendAsync(request);
                    string jsonResponse = await response.Content.ReadAsStringAsync();

                    if (jsonResponse == null)
                    {

                        return null;
                    }

                    JsonNode forecastNode = JsonNode.Parse(jsonResponse)!;

                    List<PeersList> companyPeers = JsonConvert.DeserializeObject<List<PeersList>>(jsonResponse);

                    return companyPeers;

                }
            }
        }

        public async Task<PeersProfile> GetPeersOutlook(string APIkey)
        {
            //Vai buscar a info sobre a empresa ao Financial modeling prep
            using (var httpClient = new HttpClient())

            {
                //Vai buscar as DFs
                try
                {
                    using (var request = new HttpRequestMessage(new HttpMethod("GET"), APIkey))

                    {
                        request.Headers.TryAddWithoutValidation("Upgrade-Insecure-Requests", "1");

                        var response = await httpClient.SendAsync(request);
                        string jsonResponse = await response.Content.ReadAsStringAsync();

                        if (jsonResponse == null)
                        {

                            return null;
                        }

                        JsonNode forecastNode = JsonNode.Parse(jsonResponse)!;

                        PeersProfile companyProfilesOutput = JsonConvert.DeserializeObject<PeersProfile>(jsonResponse);


                        return companyProfilesOutput;

                    }
                }
                catch (Exception)
                {

                    return new PeersProfile();
                }
                
            }
        }

        public async Task<List<MarketRiskPremium>> GetMarketRiskPremium(string APIkey)
        {
            //Vai buscar a info sobre a empresa ao Financial modeling prep
            using (var httpClient = new HttpClient())

            {
                //Vai buscar as DFs

                using (var request = new HttpRequestMessage(new HttpMethod("GET"), APIkey))

                {
                    request.Headers.TryAddWithoutValidation("Upgrade-Insecure-Requests", "1");

                    var response = await httpClient.SendAsync(request);
                    string jsonResponse = await response.Content.ReadAsStringAsync();

                    if (jsonResponse == null)
                    {

                        return null;
                    }

                    JsonNode forecastNode = JsonNode.Parse(jsonResponse)!;

                    List<MarketRiskPremium> marketRiskPremia = JsonConvert.DeserializeObject<List<MarketRiskPremium>>(jsonResponse);


                    return marketRiskPremia;

                }
            }
        }


        public async Task<List<TreasuryRates>> GetTreasuryRates(string APIkey)
        {
            //Vai buscar a info sobre a empresa ao Financial modeling prep
            using (var httpClient = new HttpClient())

            {
                //Vai buscar as DFs

                using (var request = new HttpRequestMessage(new HttpMethod("GET"), APIkey))

                {
                    request.Headers.TryAddWithoutValidation("Upgrade-Insecure-Requests", "1");

                    var response = await httpClient.SendAsync(request);
                    string jsonResponse = await response.Content.ReadAsStringAsync();

                    if (jsonResponse == null)
                    {

                        return null;
                    }

                    JsonNode forecastNode = JsonNode.Parse(jsonResponse)!;

                    List<TreasuryRates> treasuryRates = JsonConvert.DeserializeObject<List<TreasuryRates>>(jsonResponse);


                    return treasuryRates;

                }
            }
        }

        public async Task<List<CompanyNotes>> GetCompanyNotes(string APIkey)
        {
            //Vai buscar a info sobre a empresa ao Financial modeling prep
            using (var httpClient = new HttpClient())

            {
                //Vai buscar as DFs

                using (var request = new HttpRequestMessage(new HttpMethod("GET"), APIkey))

                {
                    request.Headers.TryAddWithoutValidation("Upgrade-Insecure-Requests", "1");

                    var response = await httpClient.SendAsync(request);
                    string jsonResponse = await response.Content.ReadAsStringAsync();

                    if (jsonResponse == null)
                    {

                        return null;
                    }

                    JsonNode forecastNode = JsonNode.Parse(jsonResponse)!;

                    List<CompanyNotes> companyNotes = JsonConvert.DeserializeObject<List<CompanyNotes>>(jsonResponse);


                    return companyNotes;

                }
            }
        }

        public async Task<List<StockScreener>> GetStockScreener(string APIkey)
        {
            //Vai buscar a info sobre a empresa ao Financial modeling prep
            using (var httpClient = new HttpClient())

            {
                //Vai buscar as DFs

                using (var request = new HttpRequestMessage(new HttpMethod("GET"), APIkey))

                {
                    request.Headers.TryAddWithoutValidation("Upgrade-Insecure-Requests", "1");

                    var response = await httpClient.SendAsync(request);
                    string jsonResponse = await response.Content.ReadAsStringAsync();

                    if (jsonResponse == null)
                    {

                        return null;
                    }

                    JsonNode forecastNode = JsonNode.Parse(jsonResponse)!;

                    List<StockScreener> screeners = JsonConvert.DeserializeObject<List<StockScreener>>(jsonResponse);


                    return screeners;

                }
            }
        }

        public async Task<ForexList> GetFxRate()
        {
            //Vai buscar a info sobre a empresa ao Financial modeling prep
            using (var httpClient = new HttpClient())

            {
                //Vai buscar as DFs

                using (var request = new HttpRequestMessage(new HttpMethod("GET"), "https://financialmodelingprep.com/api/v3/forex?apikey=d7ba591715e23c3d2ac3b0aec6dea138"))

                {
                    request.Headers.TryAddWithoutValidation("Upgrade-Insecure-Requests", "1");

                    var response = await httpClient.SendAsync(request);
                    string jsonResponse = await response.Content.ReadAsStringAsync();

                    if (jsonResponse == null)
                    {

                        return null;
                    }

                    JsonNode forecastNode = JsonNode.Parse(jsonResponse)!;

                    ForexList FxRate = JsonConvert.DeserializeObject<ForexList>(jsonResponse);


                    return FxRate;

                }
            }
        }

        public async Task<List<CompanyProfileFinal>> GetCompanyProfile(string API)
        {
            //Vai buscar a info sobre a empresa ao Financial modeling prep
            using (var httpClient = new HttpClient())

            {
                //Vai buscar as DFs

                using (var request = new HttpRequestMessage(new HttpMethod("GET"), API))

                {
                    request.Headers.TryAddWithoutValidation("Upgrade-Insecure-Requests", "1");

                    var response = await httpClient.SendAsync(request);
                    string jsonResponse = await response.Content.ReadAsStringAsync();

                    if (jsonResponse == null)
                    {

                        return null;
                    }

                    JsonNode forecastNode = JsonNode.Parse(jsonResponse)!;

                    List<CompanyProfileFinal> profile = JsonConvert.DeserializeObject<List<CompanyProfileFinal>>(jsonResponse);


                    return profile;

                }
            }
        }


       
    }

}
