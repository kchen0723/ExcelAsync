﻿using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using GalaSoft.MvvmLight;

namespace ExcelWvvm.ViewModel
{
    public class GoogleHistoryViewModel : ViewModelBase
    {
        public string SecurityId { get; set; }
        public DateTime StartDate { get; set; }
        public DateTime EndDate { get; set; }
        public GoogleHistoryViewModel()
        {
            this.SecurityId = "aaaa";
            this.StartDate = DateTime.Now;
            this.EndDate = this.StartDate;
        }
    }
}
