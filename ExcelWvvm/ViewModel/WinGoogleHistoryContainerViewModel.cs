using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Windows;
using GalaSoft.MvvmLight;
using GalaSoft.MvvmLight.Command;
using ExcelWvvm.Entities;
//using AutoMapper;

namespace ExcelWvvm.ViewModel
{
    public class WinGoogleHistoryContainerViewModel : ViewModelBase
    {
        private GoogleHistory m_history;
        public object CurrentViewModel { get; set; }
        public bool IsOkButtonVisible { get; set; }
        public bool IsCancelButonVisible { get; set; }

        public RelayCommand<Window> CancelCommand { get; set; }
        public RelayCommand<Window> OkCommand { get; set; }

        public WinGoogleHistoryContainerViewModel()
        {
            this.CancelCommand = new RelayCommand<Window>(OnCancel);
            this.OkCommand = new RelayCommand<Window>(OnOk);
            this.CurrentViewModel = new GoogleHistoryViewModel();
            this.IsOkButtonVisible = true;
            this.IsCancelButonVisible = true;
        }

        public void OnCancel(Window win)
        {
            if (win != null)
            {
                WindowHelper.CloseWindow(win);
            }
        }

        public void OnOk(Window win)
        {
            if (this.CurrentViewModel is GoogleHistoryViewModel)
            {
                GoogleHistoryViewModel gh = this.CurrentViewModel as GoogleHistoryViewModel;
                //Mapper.Initialize(cfg => cfg.CreateMap<GoogleHistoryViewModel, GoogleHistory>());
                //this.m_history = Mapper.Map<GoogleHistory>(gh);
                this.m_history.OnRetrievedDataHandler = History_OnRetrievedData;
                this.m_history.ExecuteAsync();
                this.CurrentViewModel = new LoadingViewModel();
                this.IsOkButtonVisible = false;
                this.RaisePropertyChanged(null);
            }
            else if (this.CurrentViewModel is LoadingViewModel)
            {
            }
            else if (this.CurrentViewModel is DataResultViewModel)
            {
                this.OnCancel(win);
                DataResultViewModel dvm = this.CurrentViewModel as DataResultViewModel;
                if (ExcelHandler.WriteToRangeHandler != null)
                {
                    ExcelHandler.WriteToRangeHandler(dvm.Result, dvm.GH);
                }
            }
        }

        private void History_OnRetrievedData(GoogleHistory gh, object[,] result)
        {
            DataResultViewModel dvm = new DataResultViewModel();
            dvm.GH = gh;
            dvm.Result = result;
            this.CurrentViewModel = dvm;
            this.IsOkButtonVisible = true;
            this.IsCancelButonVisible = false;
            this.RaisePropertyChanged(null);
        }
    }
}
