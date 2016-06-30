﻿using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using GalaSoft.MvvmLight;
using ExcelWvvm.Entities;

namespace ExcelWvvm.ViewModel
{
    public class DataResultViewModel : ViewModelBase
    {
        public GoogleHistory GH { get; set; }

        public object[,] Result { get; set; }

        public int RowCount
        {
            get
            {
                return this.Result.GetLength(0);
            }
        }

        public int ColumnCount
        {
            get
            {
                return this.Result.GetLength(1);
            }
        }
    }
}
