using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Windows;
using ExcelWvvm.Entities;

namespace ExcelWvvm
{
    public class ExcelHandler
    {
        public static Action<object[,], GoogleHistory> WriteToRangeHandler { get; set; }
    }
}
