using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;

namespace ExcelWvvm.ViewModel
{
    public class WinGoogleHistoryContainerViewModel
    {
        public GoogleHistoryViewModel GoogleHistoryVM { get; set; }

        public object CurrentViewModel { get; set; }

        public WinGoogleHistoryContainerViewModel(GoogleHistoryViewModel vm)
        {
            this.GoogleHistoryVM = vm;
        }
    }
}
