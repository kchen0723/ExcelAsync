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
    /// Interaction logic for WinDataResult.xaml
    /// </summary>
    public partial class WinDataResult : Window
    {
        private object[,] m_result;
        public Entities.GoogleHistory History { get; set; }
        public object[,] result
        {
            get
            {
                return this.m_result;
            }
            set
            {
                this.m_result = value;
                this.lblResult.Content = string.Format("{0} Rows {1} columns", m_result.GetUpperBound(0) + 1, m_result.GetUpperBound(1) + 1);
            }
        }
        public WinDataResult()
        {
            InitializeComponent();
        }

        private void button_Click(object sender, RoutedEventArgs e)
        {
            this.Close();
            if (ExcelHandler.WriteToRangeHandler != null)
            {
                this.History = ExcelHandler.WriteToRangeHandler(this.result, this.History);
            }
        }
    }
}
