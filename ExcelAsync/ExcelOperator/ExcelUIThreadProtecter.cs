using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;

namespace ExcelAsync.ExcelOperator
{
    public class ExcelUIThreadProtecter
    {
        public static void CheckIsExcelUIMainThread()
        {
            if (ExcelDna.Integration.ExcelDnaUtil.IsMainThread == false)
            {
                throw new Exception("All excel operation must be run at excel UI main thread");
            }
        }
    }
}
