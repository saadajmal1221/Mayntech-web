using System.Data;
using System.Data.SqlClient;
using System.Linq.Expressions;
using Mayntech___Individual_Solution.Auxiliar.Auxiliar.FMP;

namespace Mayntech___Individual_Solution.Auxiliar.Auxiliar.Other
{
    public class ListOfCompaniesDatabase
    {
        string str;

        public void CreateDatabse()
        {
            SqlConnection myConn = new SqlConnection("Data Source=localhost;Initial Catalog=MaynTech;Integrated Security=True");

            str = "CREATE DATABASE ListOfCompanies ON PRIMARY " +
             "(NAME = ListOfCompanies) " +
             "(symbol, name)";


            SqlCommand myCommand = new SqlCommand(str, myConn);
            try
            {
                myConn.Open();
                myCommand.ExecuteNonQuery();

            }
            catch (Exception ex)
            {

            }
            finally
            {
                if (myConn.State == ConnectionState.Open)
                {
                    myConn.Close();
                }
            }
        }

        //public void UpdateDatabase(List<StockMarketQuote> stockMarkets)
        //{
        //    try
        //    {
        //        string connectionString = "Data Source=localhost;Initial Catalog=MaynTech;Integrated Security=True";
        //        using (SqlConnection connectin = new SqlConnection(connectionString))
        //        {
        //            connectin.Open();
        //            string sql = "INSERT INTO ListOfCompanies" +
        //                "(symbol, name) VALUES" +
        //                "(@symbol, @name);";

        //            //foreach (StockMarketQuote item in stockMarkets)
        //            //{
        //            //    using (SqlCommand command = new SqlCommand(sql, connectin))
        //            //    {
        //            //        command.Parameters.AddWithValue("@symbol",);
        //            //}

        //            //}

        //        }
        //    }
        //    catch (Exception)
        //    {

        //        throw;
        //    }
        //}
    }
}
