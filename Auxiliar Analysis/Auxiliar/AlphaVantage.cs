using System.Net;
using System.Text.Json;

namespace Mayntech___Individual_Solution.Auxiliar_Analysis.Auxiliar
{
    public class AlphaVantage
    {
        public void AlphaVantageAPI()
        {
            string QUERY_URL = "https://www.alphavantage.co/query?function=INCOME_STATEMENT&symbol=SON&apikey=G43BVGDXVJIOG2QT";
            Uri queryUri = new Uri(QUERY_URL);

            using (WebClient client = new WebClient())
            {

                // -------------------------------------------------------------------------
                // if using .NET Core (System.Text.Json)
                // using .NET Core libraries to parse JSON is more complicated. For an informative blog post
                // https://devblogs.microsoft.com/dotnet/try-the-new-system-text-json-apis/

                dynamic json_data = JsonSerializer.Deserialize<Dictionary<string, dynamic>>(client.DownloadString(queryUri));

                // -------------------------------------------------------------------------

                // do something with the json_data
            }
        }
    }
}
