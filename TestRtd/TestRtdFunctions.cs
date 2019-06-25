using System;
using System.Collections.Generic;
using System.Diagnostics;
using System.Linq;
using System.Text;
using ExcelDna.Integration;

namespace CSharpAddIn
{
    public class TestRtdFunctions
    {
        [ExcelFunction(Category = "Excel-DNA RTD函数", Description = "自动刷新数据", IsMacroType = false)]
        public static string RtdArrayTest(string prefix, bool random)
        {
            string[] parm = { prefix, random.ToString() };
            object rtdValue = XlCall.RTD("CSharpAddIn.TestRtdServer", null, parm);

            string resultString = rtdValue as string;
            if (resultString == null)
                return "--";

            return resultString;
        }

        [ExcelFunction(Category = "Excel-DNA RTD函数", Description = "自动刷新数据", IsMacroType = true)]
        public static object RtdArrayTestMT(string prefix, bool random, int index)
        {
            string[] parm = { prefix, random.ToString() };
            object rtdValue = XlCall.RTD("CSharpAddIn.TestRtdServer", null, parm);

            var resultString = rtdValue as string;
            if (resultString == null)
                return "--";

            // We have a string value, parse and return as an 2x1 array
            var parts = resultString.Split(';');
            var result = new object[2, 1];
            result[0, 0] = parts[0];
            result[1, 0] = parts[1];
            return result[index,0];
        }
    }
}
