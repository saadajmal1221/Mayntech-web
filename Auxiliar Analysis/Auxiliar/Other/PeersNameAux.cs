using System.Data.SqlClient;

namespace Mayntech___Individual_Solution.Auxiliar.Auxiliar.Other
{
    public class PeersNameAux
    {
        public string GetCompanyName(string companyTick)
        {
            string companyName = null;
            try
            {
                String connection = "Data Source=localhost;Initial Catalog=MaynTech;Integrated Security=True";
                using (SqlConnection con = new SqlConnection(connection))
                {
                    con.Open();
                    String sql = "SELECT name FROM ListOfStocks WHERE symbol ='" + companyTick.Replace("'", "").Trim() + "'";
                    using (SqlCommand command = new SqlCommand(sql, con))
                    {
                        using (SqlDataReader reader = command.ExecuteReader())
                        {
                            reader.Read();
                            companyName = reader.GetString(0);
                        }
                    }

                }
            }
            catch (Exception)
			{

				throw;
			}

            return companyName;
        }
    }
}
