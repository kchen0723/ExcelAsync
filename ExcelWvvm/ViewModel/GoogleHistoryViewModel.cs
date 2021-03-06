﻿using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Windows;
using GalaSoft.MvvmLight;
using GalaSoft.MvvmLight.Command;

namespace ExcelWvvm.ViewModel
{
    public class GoogleHistoryViewModel : ViewModelBase
    {
        public string SecurityId { get; set; }
        public DateTime StartDate { get; set; }
        public DateTime EndDate { get; set; }

        public GoogleHistoryViewModel()
        {
            this.SecurityId = "GOOG";
            this.StartDate = DateTime.Now.AddDays(-7);
            this.EndDate = DateTime.Now; 
        }
    }
}
