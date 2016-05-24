using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;

namespace ExcelAsyncWvvm
{
    public class ExcelHandler
    {
        public static Func<object[,], bool> WriteToRangeHandler { get; set; }

        public static Action<ExcelDna.Integration.ExcelAction> QueueToRunUIThreadHandler { get; set; }
    }
}
