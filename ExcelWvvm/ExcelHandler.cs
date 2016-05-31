using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Windows;

namespace ExcelWvvm
{
    public class ExcelHandler
    {
        public delegate Window ShowWindowHandler(params object[] args);

        public static Func<object[,], bool> WriteToRangeHandler { get; set; }

        public static Action<ShowWindowHandler, object[]> ShowWinHandler { get; set; }
    }
}
