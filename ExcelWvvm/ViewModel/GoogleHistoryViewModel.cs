using System;
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
            this.SecurityId = "";
            this.StartDate = DateTime.Now;
            this.EndDate = this.StartDate.AddDays(7);
        }
    }
}
