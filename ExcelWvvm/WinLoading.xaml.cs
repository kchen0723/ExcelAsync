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

namespace ExcelWvvm
{
    /// <summary>
    /// Interaction logic for WinLoading.xaml
    /// </summary>
    public partial class WinLoading : Window
    {
        public event EventHandler OnCancel;
        public WinLoading()
        {
            InitializeComponent();
        }

        private void button_Click(object sender, RoutedEventArgs e)
        {
            if (this.OnCancel != null)
            {
                this.OnCancel(this, EventArgs.Empty);
            }
            this.Close();
        }
    }
}
