using Mayntech___Individual_Solution.Auxiliar.Auxiliar.Other;
using OfficeOpenXml;
using System.Drawing;

namespace Mayntech___Individual_Solution.Auxiliar.Analysis
{
    public class CreateSeparators
    {
        public async Task CreateSeparator(ExcelPackage package, string SeparatorName)
        {
            var workSheet = package.Workbook.Worksheets.Add(SeparatorName);
            workSheet.View.ShowGridLines = false;



            workSheet.TabColor = Cores.CorPrincipal;
        }
    }
}
