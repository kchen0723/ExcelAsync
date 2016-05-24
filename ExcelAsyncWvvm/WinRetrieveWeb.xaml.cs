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
using ExcelDna.Integration;

namespace ExcelAsyncWvvm
{
    /// <summary>
    /// Interaction logic for WinRetrieveWeb.xaml
    /// </summary>
    public partial class WinRetrieveWeb : Window
    {
        private Dictionary<string, string[,]> result = new Dictionary<string, string[,]>();

        public Func<string[,], bool> WriteToRangeHandler { get; set;}

        public WinRetrieveWeb()
        {
            InitializeComponent();
            this.Top = 100;
            this.Left = 100;
        }

        private void btnCancel_Click(object sender, RoutedEventArgs e)
        {
            this.Close();
        }

        private void btnOK_Click(object sender, RoutedEventArgs e)
        {
            string sitesString = new TextRange(rbSites.Document.ContentStart, rbSites.Document.ContentEnd).Text;
            string[] input = sitesString.Split(new string[] { "\r", "\n" }, StringSplitOptions.RemoveEmptyEntries);
            if (input != null && input.Length > 0)
            {
                string[,] item = new string[input.Length, 2];
                for (int i = 0; i < input.Length; i++)
                {
                    item[i, 0] = (i * i).ToString();
                    item[i, 1] = input[i];
                }
                result.Add("test1", item);
                ExcelAsyncUtil.QueueAsMacro(postToExcel);
            }
        }

        private void postToExcel()
        {
            if (result.Count > 0)
            {
                if (result.ContainsKey("test1"))
                {
                    string[,] response = result["test1"];
                    //dynamic xlApp = ExcelDnaUtil.Application;
                    //dynamic sheet2 = xlApp.ActiveWorkbook.WorkSheets("Sheet2");
                    //dynamic newSheet = xlApp.ActiveWorkbook.Worksheets.Add(After: sheet2);
                    //dynamic range = newSheet.Range("A1:B" + (response.GetUpperBound(0) + 1).ToString());
                    //range.Value = response;
                    //ExcelOperator.ReadWriteRange.WriteToRange(response);
                    if(this.WriteToRangeHandler != null)
                    {
                        this.WriteToRangeHandler(response);
                    }
                }
                result.Remove("test1");
            }
        }
    }
}
