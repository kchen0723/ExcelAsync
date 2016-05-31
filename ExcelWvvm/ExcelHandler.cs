using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Windows;

namespace ExcelWvvm
{
    public class ExcelHandler
    {
        public static Func<object[,], bool> WriteToRangeHandler { get; set; }
    }
}
