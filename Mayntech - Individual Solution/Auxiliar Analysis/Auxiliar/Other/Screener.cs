using Mayntech___Individual_Solution.Auxiliar.Auxiliar.FMP;
using System.Data.SqlClient;

namespace Mayntech___Individual_Solution.Auxiliar_Analysis.Auxiliar.Other
{
    public class Screener
    {
        public async void StockScreener()
        {
            List<string> symbols = new List<string>();

            String connection = "Data Source=localhost;Initial Catalog=MaynTech;Integrated Security=True";
            using (SqlConnection conn = new SqlConnection(connection))
            {
                conn.Open();

                string select = "SELECT [symbol] FROM ListOfStocks";

                using (SqlCommand cmd = new SqlCommand(select, conn))
                {

                    using (SqlDataReader reader = cmd.ExecuteReader())
                    {
                        while (reader.Read())
                        {

                            symbols.Add(reader.GetString(0));
                        }

                    }
                    conn.Close();
                }
            }

            for (int i = 0; i < symbols.Count(); i++)
            {
                string APICompanyOutlook = "https://financialmodelingprep.com/api/v4/company-outlook?symbol=" + symbols[i] + "&apikey=d7ba591715e23c3d2ac3b0aec6dea138";

                FinancialMP fmp = new FinancialMP();

                CompanyProfile companyOutlook = await fmp.GetCompanyOutlook(APICompanyOutlook);

                string connectionString = "Data Source=localhost;Initial Catalog=MaynTech;Integrated Security=True";

                using (SqlConnection connection2 = new SqlConnection(connectionString))
                {
                    connection2.Open();

                    string sqlcompanyOutlook = "INSERT INTO companyOutlook" +
                    "(symbol, name, industry," +
                    " sector, country) VALUES" +
                    "(@symbol,  @name, @industry" +
                    ", @sector, @country);";

                    using (SqlCommand command = new SqlCommand(sqlcompanyOutlook, connection2))
                    {
                        try
                        {
                            command.Parameters.AddWithValue("@symbol", companyOutlook.profile.symbol);
                            command.Parameters.AddWithValue("@name", companyOutlook.profile.companyName.Replace("'", "").Trim());
                            command.Parameters.AddWithValue("@sector", companyOutlook.profile.sector.Replace("'", "").Trim());
                            command.Parameters.AddWithValue("@industry", companyOutlook.profile.industry.Replace("'", "").Trim());
                            command.Parameters.AddWithValue("@country", companyOutlook.profile.country.Replace("'", "").Trim());

                            command.ExecuteNonQuery();
                        }
                        catch
                        {
                            continue;
                        }


                    }
                }

            }
        }

        public void NumberOfIndustries()
        {
            List<string> industry = new List<string>();
            String connection = "Data Source=localhost;Initial Catalog=MaynTech;Integrated Security=True";
            using (SqlConnection conn = new SqlConnection(connection))
            {
                conn.Open();

                string select = "SELECT DISTINCT [industry] FROM companyOutlook";

                using (SqlCommand cmd = new SqlCommand(select, conn))
                {

                    using (SqlDataReader reader = cmd.ExecuteReader())
                    {
                        while (reader.Read())
                        {

                            industry.Add(reader.GetString(0));
                        }

                    }
                    conn.Close();
                }
            }
        }

        public void NumberOfSectors()
        {
            List<string> sector = new List<string>();
            String connection = "Data Source=localhost;Initial Catalog=MaynTech;Integrated Security=True";
            using (SqlConnection conn = new SqlConnection(connection))
            {
                conn.Open();

                string select = "SELECT DISTINCT [sector] FROM companyOutlook";

                using (SqlCommand cmd = new SqlCommand(select, conn))
                {

                    using (SqlDataReader reader = cmd.ExecuteReader())
                    {
                        while (reader.Read())
                        {

                            sector.Add(reader.GetString(0));
                        }

                    }
                    conn.Close();
                }
            }
        }
    }
}
