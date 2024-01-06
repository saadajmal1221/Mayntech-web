using Mayntech___Individual_Solution.Auxiliar.Auxiliar.FMP;
using Microsoft.AspNetCore.DataProtection.KeyManagement;
using Newtonsoft.Json;
using System.Data.SqlClient;
using System.Globalization;
using System.Text.Json.Nodes;

namespace Mayntech___Individual_Solution.Auxiliar_Analysis.Auxiliar.Other
{
    public class DatabaseUpdate
    {
        public List<StockScreener> screenersTech = new List<StockScreener>();
        public List<StockScreener> screenersEnergy = new List<StockScreener>();
        public List<StockScreener> screenersConsCyclic = new List<StockScreener>();
        public List<StockScreener> screenersConsDefe = new List<StockScreener>();
        public List<StockScreener> screenersIndustrial = new List<StockScreener>();
        public List<StockScreener> screenersHealth = new List<StockScreener>();
        public List<StockScreener> screenersRealEsta = new List<StockScreener>();
        public List<StockScreener> screenersCommun = new List<StockScreener>();
        public List<StockScreener> screenersconglomera = new List<StockScreener>();
        public List<StockScreener> screenersMater = new List<StockScreener>();
        public List<StockScreener> screenersutilities = new List<StockScreener>();
        public List<List<StockScreener>> AllStocksData = new List<List<StockScreener>>();

        public async void Database()
        {
            FinancialMP fmp = new FinancialMP();

            string APISOtckScreenerTechnology = "https://financialmodelingprep.com/api/v3/stock-screener?isEtf=false&isActivelyTrading=true&limit=99999&&sector=Technology&apikey=d7ba591715e23c3d2ac3b0aec6dea138";
            screenersTech = await fmp.GetStockScreener(APISOtckScreenerTechnology);
            AllStocksData.Add(screenersTech);

            string APISOtckScreenerEnergy = "https://financialmodelingprep.com/api/v3/stock-screener?isEtf=false&isActivelyTrading=true&limit=99999&&sector=Energy&apikey=d7ba591715e23c3d2ac3b0aec6dea138";
            screenersEnergy = await fmp.GetStockScreener(APISOtckScreenerEnergy);
            AllStocksData.Add(screenersEnergy);

            string APISOtckScreenerConsumerCyclical = "https://financialmodelingprep.com/api/v3/stock-screener?isEtf=false&isActivelyTrading=true&limit=99999&&sector=Consumer%20Cyclical&apikey=d7ba591715e23c3d2ac3b0aec6dea138";
            screenersConsCyclic = await fmp.GetStockScreener(APISOtckScreenerConsumerCyclical);
            AllStocksData.Add(screenersConsCyclic);

            string APISOtckScreenerIndustrials = "https://financialmodelingprep.com/api/v3/stock-screener?isEtf=false&isActivelyTrading=true&limit=99999&&sector=Industrials&apikey=d7ba591715e23c3d2ac3b0aec6dea138";
            screenersIndustrial = await fmp.GetStockScreener(APISOtckScreenerIndustrials);
            AllStocksData.Add(screenersIndustrial);

            string APISOtckScreenerMaterials = "https://financialmodelingprep.com/api/v3/stock-screener?isEtf=false&isActivelyTrading=true&limit=99999&&sector=Materials&apikey=d7ba591715e23c3d2ac3b0aec6dea138";
            screenersMater = await fmp.GetStockScreener(APISOtckScreenerMaterials);
            AllStocksData.Add(screenersMater);

            string APISOtckScreenerCommunication = "https://financialmodelingprep.com/api/v3/stock-screener?isEtf=false&isActivelyTrading=true&limit=99999&&sector=Communication&apikey=d7ba591715e23c3d2ac3b0aec6dea138";
            screenersCommun = await fmp.GetStockScreener(APISOtckScreenerCommunication);
            AllStocksData.Add(screenersCommun);

            string APISOtckScreenerConsumerDefensive = "https://financialmodelingprep.com/api/v3/stock-screener?isEtf=false&isActivelyTrading=true&limit=99999&&sector=Defensive&apikey=d7ba591715e23c3d2ac3b0aec6dea138";
            screenersConsDefe = await fmp.GetStockScreener(APISOtckScreenerConsumerDefensive);
            AllStocksData.Add(screenersConsDefe);

            string APISOtckScreenerHealth = "https://financialmodelingprep.com/api/v3/stock-screener?isEtf=false&isActivelyTrading=true&limit=99999&&sector=Healthcare&apikey=d7ba591715e23c3d2ac3b0aec6dea138";
            screenersHealth = await fmp.GetStockScreener(APISOtckScreenerHealth);
            AllStocksData.Add(screenersHealth);

            string APISOtckScreenerrealEstate = "https://financialmodelingprep.com/api/v3/stock-screener?isEtf=false&isActivelyTrading=true&limit=99999&&sector=real%20Estate&apikey=d7ba591715e23c3d2ac3b0aec6dea138";
            screenersRealEsta = await fmp.GetStockScreener(APISOtckScreenerrealEstate);
            AllStocksData.Add(screenersRealEsta);

            string APISOtckScreenerUtilities = "https://financialmodelingprep.com/api/v3/stock-screener?isEtf=false&isActivelyTrading=true&limit=99999&&sector=Utilities&apikey=d7ba591715e23c3d2ac3b0aec6dea138";
            screenersutilities = await fmp.GetStockScreener(APISOtckScreenerUtilities);
            AllStocksData.Add(screenersutilities);

            string APISOtckScreenerconglomerates = "https://financialmodelingprep.com/api/v3/stock-screener?isEtf=false&isActivelyTrading=true&limit=99999&&sector=Conglomerate&apikey=d7ba591715e23c3d2ac3b0aec6dea138";
            screenersconglomera = await fmp.GetStockScreener(APISOtckScreenerconglomerates);
            AllStocksData.Add(screenersconglomera);

            List<StockScreener> AllStocks = new List<StockScreener>();
            for (int i = 0; i < AllStocksData.Count(); i++)
            {
                foreach (StockScreener item in AllStocksData[i])
                {

                    if (AllStocks.Any(screener => screener.companyName == item.companyName))
                    {
                        StockScreener? objScreen = AllStocks.Find(x => (x.companyName == item.companyName));

                        try
                        {
                            if (objScreen.symbol.Length <= item.symbol.Length)
                            {

                            }
                            else
                            {
                                objScreen.symbol = item.symbol;
                            }
                        }
                        catch
                        {

                            continue;
                        }

                    }
                    else
                    {
                        AllStocks.Add(item);
                    }

                }
            }


            try
            {
                string connectionString = "Data Source=localhost;Initial Catalog=MaynTech;Integrated Security=True";

                using (SqlConnection connection = new SqlConnection(connectionString))
                {

                    connection.Open();

                    //Apaga tudo o que temos
                    string delete = "DELETE FROM companyInfo";
                    string delete2 = "DELETE FROM ListOfStocks";

                    using (SqlCommand c = new SqlCommand(delete, connection))
                    {
                        c.ExecuteNonQuery();
                    }

                    using (SqlCommand c = new SqlCommand(delete2, connection))
                    {
                        c.ExecuteNonQuery();
                    }

                    //Cria uma nova ListOfStocks

                    string sqlListOfStocks = "INSERT INTO ListOfStocks" +
                        "(symbol, name) VALUES" + "(@symbol, @name);";

                    foreach (var item in AllStocks)
                    {
                        try
                        {
                            using (SqlCommand command = new SqlCommand(sqlListOfStocks, connection))
                            {
                                try
                                {
                                    command.Parameters.AddWithValue("@symbol", item.symbol);
                                    command.Parameters.AddWithValue("@name", item.companyName.Replace("'", "").Trim());

                                    command.ExecuteNonQuery();
                                }
                                catch
                                {
                                    continue;
                                }

                            }


                        }
                        catch (Exception)
                        {


                        }


                    }
                }
            }
            catch (Exception ex)
            {

            }
        }

    }
}
