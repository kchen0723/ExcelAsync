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
            WpfWindowHelper.ShowWindow<WinLoading>(getLoadingInstance);
        }

        private void History_OnRetrievedData(object sender, EventArgs e)
        {
            WpfWindowHelper.CloseWindow(this.loadingWindow);
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
