using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Linq;
using System.Runtime.InteropServices;
using System.Text;
using System.Windows.Forms;

using ExcelDna.Integration;

namespace ExcelAsyncWinForm
{
    public partial class FrmRetrieveWeb : Form
    {
        [DllImport("user32.dll")]
        public static extern IntPtr SetParent(IntPtr hWndChild, IntPtr hWndNewParent);

        private Dictionary<string, string[,]> result = new Dictionary<string, string[,]>();

        public FrmRetrieveWeb()
        {
            InitializeComponent();
            SetParent(this.Handle, ExcelDnaUtil.WindowHandle);
        }

        private void btnCancel_Click(object sender, EventArgs e)
        {
            this.Close();
        }

        private void btnOK_Click(object sender, EventArgs e)
        {
            string[] input = this.rbSites.Text.Split(new string[] { "\r", "\n" }, StringSplitOptions.RemoveEmptyEntries);
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
                    dynamic xlApp = ExcelDnaUtil.Application;
                    dynamic sheet2 = xlApp.ActiveWorkbook.WorkSheets("Sheet2");
                    dynamic newSheet = xlApp.ActiveWorkbook.Worksheets.Add(After: sheet2);
                    dynamic range = newSheet.Range("A1:B" + (response.GetUpperBound(0) + 1).ToString());
                    range.Value = response;
                }
                result.Remove("test1");
            }
        }
    }
}
