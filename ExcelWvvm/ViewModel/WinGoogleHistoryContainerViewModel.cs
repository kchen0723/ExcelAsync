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
        public object CurrentViewModel { get; set; }

        public bool IsOkButtonVisibal { get; set; }
        public bool IsCancelButonVisibal { get; set; }

        public RelayCommand<Window> CancelCommand { get; set; }

        public RelayCommand OkCommand { get; set; }

        public WinGoogleHistoryContainerViewModel()
        {
            this.CancelCommand = new RelayCommand<Window>(OnCancel);
            this.OkCommand = new RelayCommand(OnOk);
            this.CurrentViewModel = new GoogleHistoryViewModel();
            this.IsOkButtonVisibal = true;
            this.IsCancelButonVisibal = true;
        }

        public void OnCancel(Window win)
        {
            if (win != null)
            {
                WindowHelper.CloseWindow(win);
            }
        }

        public void OnOk()
        {
            if (this.CurrentViewModel is GoogleHistoryViewModel)
            {
                GoogleHistoryViewModel gh = this.CurrentViewModel as GoogleHistoryViewModel;
                this.CurrentViewModel = new LoadingViewModel();
                //this.IsOkButtonVisibal = false;
                this.RaisePropertyChanged("CurrentViewModel");
                this.RaisePropertyChanged("IsOkButtonVisibal");
            }
            else if (this.CurrentViewModel is LoadingViewModel)
            {
                this.CurrentViewModel = new DataResultViewModel();
                this.RaisePropertyChanged("CurrentViewModel");
            }
        }
    }
}
