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
        GoogleHistory history = null;
        public WinGoogleHistory()
        {
            InitializeComponent();
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
            history.OnRetrievedDataHandler = History_OnRetrievedData;
            history.ExecuteAsync();
            WindowHelper.ShowWindow(createLoadingInstance, null);
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
            WindowHelper.ShowWindow<WinDataResult>(showDataResult, new object[] { arg2 as object[,], this.history });
        }

        private void showDataResult(Window win, params object[] args)
        {
            WinDataResult resultWin = win as WinDataResult;
            if (args != null && args.Length == 2)
            {
                resultWin.result = args[0] as object[,];
                resultWin.History = args[1] as GoogleHistory;
            }
        }

        private void LoadingWindow_OnCancel(object sender, EventArgs e)
        {
            history.CancelExecute();
        }
    }
}
