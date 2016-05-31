using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Windows;
using System.Windows.Controls;
using System.Windows.Data;
using System.Windows.Documents;
using System.Windows.Input;
using System.Windows.Media;
using System.Windows.Media.Imaging;
using System.Windows.Shapes;
using ExcelWvvm.Entities;

namespace ExcelWvvm
{
    /// <summary>
    /// Interaction logic for WinGoogleHistory.xaml
    /// </summary>
    public partial class WinGoogleHistory : Window
    {
        WinLoading loadingWindow = null;
        WinDataResult resultWin = null;
        GoogleHistory history = null;
        object[,] result = null;
        public WinGoogleHistory()
        {
            InitializeComponent();
            this.Left = 200;
            this.Top = 200;
        }

        private void btnCancel_Click(object sender, RoutedEventArgs e)
        {
            this.Close();
        }

        private void btnOK_Click(object sender, RoutedEventArgs e)
        {
            history = new GoogleHistory();
            history.SecurityId = this.tbSecurityId.Text;
            history.StartDate = DateTime.Parse(this.tbStartDate.Text);
            history.EndDate = DateTime.Parse(this.tbEndDate.Text);

            this.Visibility = Visibility.Hidden;
            history.OnRetrievedData += History_OnRetrievedData;
            history.ExecuteAsync();
            //WpfWindowHelper.ShowWindow<WinLoading>(getLoadingInstance);
            if (ExcelHandler.ShowWinHandler != null)
            {
                ExcelHandler.ShowWinHandler(createLoadingInstance, null);
            }
        }

        private Window createLoadingInstance(params object[] args)
        {
            this.loadingWindow = new WinLoading();
            this.loadingWindow.OnCancel += LoadingWindow_OnCancel;
            return this.loadingWindow;
        }

        private void History_OnRetrievedData(object arg1, object arg2)
        {
            WindowHelper.CloseWindow(this.loadingWindow);
            result = arg2 as object[,];
            WindowHelper.ShowWindow<WinDataResult>(showDataResult);
        }

        private void showDataResult(object sender, EventArgs e)
        {
            this.resultWin = sender as WinDataResult;
            this.resultWin.result = this.result;
        }

        private void getLoadingInstance(object sender, EventArgs e)
        {
            this.loadingWindow = sender as WinLoading;
            this.loadingWindow.OnCancel += LoadingWindow_OnCancel;
        }

        private void LoadingWindow_OnCancel(object sender, EventArgs e)
        {
            history.CancelExecute();
        }
    }
}
