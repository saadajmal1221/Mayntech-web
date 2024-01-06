using Mayntech___Individual_Solution.Auxiliar.Auxiliar.FMP;
using System.Collections.Generic;

namespace Mayntech___Individual_Solution.Auxiliar_Analysis.Auxiliar.Other
{
    public class NormalizePeersStatements
    {
        public List<FinancialStatements> FinancialStatementsNormalizations(PeersProfile peersOutlook, DateTime ReferenceDate, int numberOfYears)
        {
            try
            {
                DateTime date = peersOutlook.financialsAnnual.income[0].Date;

                if (peersOutlook.profile.isActivelyTrading == false)
                {
                    return null;
                }

                int quarters = quartersDiff(ReferenceDate, date);

                if (quarters>4)
                {
                    return null;
                }

                List<FinancialStatements> ListAux = Normalization(quarters, peersOutlook, numberOfYears);


                return ListAux;
            }
            catch 
            {

                return null;
            }

        }

        public int quartersDiff(DateTime referencedate, DateTime date)
        {
            int aux = DateTime.Compare(date, referencedate);


            int output = 0;
            if (aux > 0)
            {
                output = DateGreaterThanReference(referencedate, date);
            }
            else if (aux < 0)
            {
                output = DateSmallerThanReference(referencedate, date);
            }
            return output;

        }

        public int DateGreaterThanReference(DateTime referencedate, DateTime date)
        {
            int output = 0;
            date = date.AddMonths(-3);

            int aux = DateTime.Compare(date, referencedate);

            if (aux<=0)
            {
                if (date.Month < referencedate.Month-2)
                {
                    return output;
                }
                else
                {
                    output -= 1;
                    return output;
                }

            }
            else
            {
                output -= 1;

                output += DateGreaterThanReference(referencedate, date);
                return output;
            }
            
        }

        public int DateSmallerThanReference(DateTime referencedate, DateTime date)
        {
            int output = 0;
            date = date.AddMonths(3);

            int aux = DateTime.Compare(date, referencedate);

            if (aux >= 0)
            {
                if (date.Month > referencedate.Month + 2)
                {
                    return output;
                }
                else
                {
                    output += 1;
                    return output;
                }

            }
            else
            {
                output += 1;

                output += DateSmallerThanReference(referencedate, date);
                return output;
            }
            
        }
        public List<FinancialStatements> Normalization(int quarters , PeersProfile peersOutlook, int numberOfYears)
        {
            int numberOfPeriods = Math.Min(numberOfYears, 5);

            List<FinancialStatements> financials = new List<FinancialStatements>();

            if (peersOutlook.financialsAnnual.cash.Count()>=numberOfPeriods && peersOutlook.financialsAnnual.balance.Count() >= numberOfPeriods
                && peersOutlook.financialsAnnual.income.Count() >= numberOfPeriods)
            {
                double multiplicationFirstAux = ((double)4 - (double)Math.Abs(quarters)) / (double)4;
                double multiplicationSecondAux = (double)Math.Abs(quarters) / (double)4;

                for (int i = 0; i < numberOfPeriods; i++)
                {
                    FinancialStatements financialStatementsAux = new FinancialStatements();
                    int aux = i + 1;

                    if (quarters<0)
                    {
                        if (i == numberOfPeriods - 1)
                        {
                            aux = i;
                        }
                        else
                        {
                            aux = i + 1;
                        }
                        
                    }

                    if (quarters > 0)
                    {
                        if (i == 0)
                        {
                            aux = i;
                        }
                        else
                        {
                            aux = i - 1;
                        }
                        
                    }
                    if (quarters == 0)
                    {
                        aux = i;
                    }


                    //if (i==0 || i==numberOfPeriods-1)
                    //{
                    //    aux = i;
                    //}


                    //BALANCE
                    //Receivables
                    financialStatementsAux.NetReceivables = peersOutlook.financialsAnnual.balance[i].NetReceivables * multiplicationFirstAux
                        + peersOutlook.financialsAnnual.balance[aux].NetReceivables * multiplicationSecondAux;


                    //Payables
                    financialStatementsAux.AccountPayables = peersOutlook.financialsAnnual.balance[i].AccountPayables * multiplicationFirstAux
    + peersOutlook.financialsAnnual.balance[aux].AccountPayables * multiplicationSecondAux;


                    //current assets
                    financialStatementsAux.TotalCurrentAssets = peersOutlook.financialsAnnual.balance[i].TotalCurrentAssets * multiplicationFirstAux
+ peersOutlook.financialsAnnual.balance[aux].TotalCurrentAssets * multiplicationSecondAux;

                    //current Liabilities
                    financialStatementsAux.totalCurrentLiabilities = peersOutlook.financialsAnnual.balance[i].totalCurrentLiabilities * multiplicationFirstAux
+ peersOutlook.financialsAnnual.balance[aux].totalCurrentLiabilities * multiplicationSecondAux;

                    //cash
                    financialStatementsAux.CashAndCashEquivalents = peersOutlook.financialsAnnual.balance[i].CashAndCashEquivalents * multiplicationFirstAux
+ peersOutlook.financialsAnnual.balance[aux].CashAndCashEquivalents * multiplicationSecondAux;

                    //Shor term invemtnes
                    financialStatementsAux.ShortTermInvestments = peersOutlook.financialsAnnual.balance[i].ShortTermInvestments * multiplicationFirstAux
+ peersOutlook.financialsAnnual.balance[aux].ShortTermInvestments * multiplicationSecondAux;

                    //PP&E
                    financialStatementsAux.propertyPlantEquipmentNet = peersOutlook.financialsAnnual.balance[i].propertyPlantEquipmentNet * multiplicationFirstAux
+ peersOutlook.financialsAnnual.balance[aux].propertyPlantEquipmentNet * multiplicationSecondAux;

                    //Inventory
                    financialStatementsAux.Inventory = peersOutlook.financialsAnnual.balance[i].Inventory * multiplicationFirstAux
+ peersOutlook.financialsAnnual.balance[aux].Inventory * multiplicationSecondAux;

                    //Goodwill
                    financialStatementsAux.Goodwill = peersOutlook.financialsAnnual.balance[i].Goodwill * multiplicationFirstAux
+ peersOutlook.financialsAnnual.balance[aux].Goodwill * multiplicationSecondAux;

                    //Intangible assets
                    financialStatementsAux.IntangibleAssets = peersOutlook.financialsAnnual.balance[i].IntangibleAssets * multiplicationFirstAux
+ peersOutlook.financialsAnnual.balance[aux].IntangibleAssets * multiplicationSecondAux;

                    //Long term investments
                    financialStatementsAux.LongtermInvestments = peersOutlook.financialsAnnual.balance[i].LongtermInvestments * multiplicationFirstAux
+ peersOutlook.financialsAnnual.balance[aux].LongtermInvestments * multiplicationSecondAux;

                    //Other liabilities
                    financialStatementsAux.OtherLiabilities = peersOutlook.financialsAnnual.balance[i].OtherLiabilities * multiplicationFirstAux
+ peersOutlook.financialsAnnual.balance[aux].OtherLiabilities * multiplicationSecondAux;

                    //total liabilities
                    financialStatementsAux.totalLiabilities = peersOutlook.financialsAnnual.balance[i].totalLiabilities * multiplicationFirstAux
+ peersOutlook.financialsAnnual.balance[aux].totalLiabilities * multiplicationSecondAux;

                    //total debt
                    financialStatementsAux.TotalDebt = peersOutlook.financialsAnnual.balance[i].TotalDebt * multiplicationFirstAux
+ peersOutlook.financialsAnnual.balance[aux].TotalDebt * multiplicationSecondAux;

                    //total equity
                    financialStatementsAux.TotalEquity = peersOutlook.financialsAnnual.balance[i].TotalEquity * multiplicationFirstAux
+ peersOutlook.financialsAnnual.balance[aux].TotalEquity * multiplicationSecondAux;

                    //total assets
                    financialStatementsAux.TotalAssets = peersOutlook.financialsAnnual.balance[i].TotalAssets * multiplicationFirstAux
+ peersOutlook.financialsAnnual.balance[aux].TotalAssets * multiplicationSecondAux;



                    //P&L
                    //Revenue
                    financialStatementsAux.Revenue = peersOutlook.financialsAnnual.income[i].Revenue * multiplicationFirstAux
+ peersOutlook.financialsAnnual.income[aux].Revenue * multiplicationSecondAux;

                    //cost of Revenue
                    financialStatementsAux.CostOfRevenue = peersOutlook.financialsAnnual.income[i].CostOfRevenue * multiplicationFirstAux
+ peersOutlook.financialsAnnual.income[aux].CostOfRevenue * multiplicationSecondAux;

                    //Gross Profit 
                    financialStatementsAux.GrossProfit = peersOutlook.financialsAnnual.income[i].GrossProfit * multiplicationFirstAux
+ peersOutlook.financialsAnnual.income[aux].GrossProfit * multiplicationSecondAux;

                    //Gross Profit Ratio
                    financialStatementsAux.GrossProfitRatio = peersOutlook.financialsAnnual.income[i].GrossProfitRatio * multiplicationFirstAux
+ peersOutlook.financialsAnnual.income[aux].GrossProfitRatio * multiplicationSecondAux;

                    //Operatin Profit 
                    financialStatementsAux.OperatingIncome = peersOutlook.financialsAnnual.income[i].OperatingIncome * multiplicationFirstAux
+ peersOutlook.financialsAnnual.income[aux].OperatingIncome * multiplicationSecondAux;

                    //Operatin Profit ratio
                    financialStatementsAux.OperatingIncomeRatio = peersOutlook.financialsAnnual.income[i].OperatingIncomeRatio * multiplicationFirstAux
+ peersOutlook.financialsAnnual.income[aux].OperatingIncomeRatio * multiplicationSecondAux;

                    //EBITDA
                    financialStatementsAux.EBITDA = peersOutlook.financialsAnnual.income[i].EBITDA * multiplicationFirstAux
+ peersOutlook.financialsAnnual.income[aux].EBITDA * multiplicationSecondAux;

                    //EBITDA ratio
                    financialStatementsAux.EBITDARatio = peersOutlook.financialsAnnual.income[i].EBITDARatio * multiplicationFirstAux
+ peersOutlook.financialsAnnual.income[aux].EBITDARatio * multiplicationSecondAux;

                    //NetIncome
                    financialStatementsAux.NetIncome = peersOutlook.financialsAnnual.income[i].NetIncome * multiplicationFirstAux
+ peersOutlook.financialsAnnual.income[aux].NetIncome * multiplicationSecondAux;

                    //NetIncomeRatio
                    financialStatementsAux.NetIncomeRatio = peersOutlook.financialsAnnual.income[i].NetIncomeRatio * multiplicationFirstAux
+ peersOutlook.financialsAnnual.income[aux].NetIncomeRatio * multiplicationSecondAux;



                    //CFS
                    //Cash from operations
                    financialStatementsAux.NetCashProvidedByOperatingActivities = peersOutlook.financialsAnnual.cash[i].NetCashProvidedByOperatingActivities * multiplicationFirstAux
+ peersOutlook.financialsAnnual.cash[aux].NetCashProvidedByOperatingActivities * multiplicationSecondAux;

                    //Cash from financing
                    financialStatementsAux.NetCashUsedProvidedByFinancingActivities = peersOutlook.financialsAnnual.cash[i].NetCashUsedProvidedByFinancingActivities * multiplicationFirstAux
+ peersOutlook.financialsAnnual.cash[aux].NetCashUsedProvidedByFinancingActivities * multiplicationSecondAux;

                    //Cash from investing
                    financialStatementsAux.netCashUsedForInvestingActivites = peersOutlook.financialsAnnual.cash[i].netCashUsedForInvestingActivites * multiplicationFirstAux
+ peersOutlook.financialsAnnual.cash[aux].netCashUsedForInvestingActivites * multiplicationSecondAux;

                    financials.Add(financialStatementsAux);
                }


            }



            return financials;
        }
    }
}
