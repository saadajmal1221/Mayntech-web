using Mayntech___Individual_Solution.Auxiliar.Auxiliar.FMP;

namespace Mayntech___Individual_Solution.Auxiliar_Analysis.CompetitorsRatios
{
    public class CompetitorsList
    {
        public List<string> GetCompetitorsList(List<PeersList> companypeers)
        {
            List<string> competitors = new List<string>();
            foreach (PeersList item in companypeers)
            {
                foreach (string competitor in item.peersList)
                {
                    competitors.Add(competitor);
                }
            }
            return competitors;

        }
    }
}
