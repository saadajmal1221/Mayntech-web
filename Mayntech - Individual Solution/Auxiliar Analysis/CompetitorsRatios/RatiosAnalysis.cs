using Mayntech___Individual_Solution.Auxiliar.Auxiliar.FMP;
using Mayntech___Individual_Solution.Auxiliar.Auxiliar.Other;
using Mayntech___Individual_Solution.Pages;
using Microsoft.AspNetCore.Mvc.ViewFeatures;
using OfficeOpenXml;
using OfficeOpenXml.ConditionalFormatting;
using OfficeOpenXml.Drawing.Chart;
using OfficeOpenXml.Style;
using System.Drawing;
using System.Globalization;
using System.Reflection;
using static OfficeOpenXml.ExcelErrorValue;

namespace Mayntech___Individual_Solution.Auxiliar_Analysis.CompetitorsRatios
{
    public class RatiosAnalysis : CommentsRatios
    {

        public async void RatioConstruction(ExcelPackage package, int numberOfYears, IDictionary<string,
            List<FinancialStatements>> incomeStatementPeers, IDictionary<string, List<FinancialStatements>> BalanceSheetPeers,
            IDictionary<string, List<FinancialStatements>> CashFlowPeers, int numberOfCompetitors,
            List<FinancialStatements> incomeStatement, List<FinancialStatements> balanceSheet,
            List<FinancialStatements> cashFlow, string companyTick, int numberOfYearsIncomeStatement)
        {
            List<double> points = new List<double>();
            Ratios ratios = new Ratios();

            Liquidity liquidity = new Liquidity();

            List<string> commentsLiquidity = new List<string>{ "Comentário 1", "Comentário 2", "comentário 3" };

            liquidity.LiquidityConstruction(package, numberOfYears, commentsLiquidity, companyTick, numberOfYearsIncomeStatement);


            List<string> commentsSolvency = new List<string> { "Comentário 1", "Comentário 2", "comentário 3" };
            Solvency solvency = new Solvency();
            solvency.SolvencyConstruction(package, numberOfYears, commentsSolvency, companyTick, numberOfYearsIncomeStatement);

            List<string> commentsactivity = new List<string> { "Comentário 1", "Comentário 2", "comentário 3" };
            Activity activity = new Activity();
            activity.ActivityConstruction(package, numberOfYears, commentsSolvency, companyTick, numberOfYearsIncomeStatement);

            List<string> CommentsProfitability = new List<string> { "Comentário 1", "Comentário 2", "comentário 3" };
            Profitability profitability = new Profitability();
            profitability.ProfitabilityConstruction(package, numberOfYears, commentsSolvency, companyTick, numberOfYearsIncomeStatement);

        }



    }
}
