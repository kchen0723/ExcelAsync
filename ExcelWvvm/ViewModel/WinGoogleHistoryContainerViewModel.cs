using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Windows;
using GalaSoft.MvvmLight;
using GalaSoft.MvvmLight.Command;

namespace ExcelWvvm.ViewModel
{
    public class WinGoogleHistoryContainerViewModel : ViewModelBase
    {
        public GoogleHistoryViewModel GoogleHistoryVM { get; set; }

        public object CurrentViewModel { get; set; }

        public RelayCommand<Window> CloseCommand { get; set; }

        public WinGoogleHistoryContainerViewModel()
        {
            this.CloseCommand = new RelayCommand<Window>(CloseWindow);
            this.CurrentViewModel = new GoogleHistoryViewModel();
        }

        public void CloseWindow(Window win)
        {
            if (win != null)
            {
                WindowHelper.CloseWindow(win);
            }
        }
    }
}
