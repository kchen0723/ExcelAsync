using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading;

namespace ExcelAsync.ExcelOperator
{
    public class ExcelUIThreadProtecter
    {
        public static void CheckIsExcelUIMainThread()
        {
            //In practice the managed Main thread id is 1.
            if (Thread.CurrentThread.ManagedThreadId != 1)
            {
                throw new Exception("All excel operation must be run at excel UI main thread");
            }
        }
    }
}
