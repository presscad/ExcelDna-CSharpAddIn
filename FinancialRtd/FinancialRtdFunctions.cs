using System;
using System.Collections.Generic;
using System.Diagnostics;
using System.Linq;
using System.Text;
using ExcelDna.Integration;

namespace CSharpAddIn
{
    public class FinancialRtdFunctions
    {
        [ExcelFunction(Category = "Excel-DNA RTD函数", Description = "自动刷新股票数据", IsMacroType = false)]
        public static string RtdArrayFinancial(string code, string info)
        {
            string[] parm = { code , info };
            object rtdValue = XlCall.RTD("CSharpAddIn.FinancialRtdServer", null, parm);

            string resultString = rtdValue as string;
            if (resultString == null)
                return "--";

            return resultString;
        }
    }
}
