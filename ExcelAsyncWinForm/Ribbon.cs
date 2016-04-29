using System;
using System.Collections.Generic;
using System.Linq;
using System.Reflection;
using System.Runtime.InteropServices;
using System.Text;
using System.Xml;
using System.Windows.Forms;

using ExcelDna.Integration;
using ExcelDna.Integration.CustomUI;

namespace ExcelAsyncWinForm
{
    [ComVisible(true)]
    public class Ribbon : ExcelRibbon
    {
        public override string GetCustomUI(string RibbonID)
        {
            XmlDocument doc = new XmlDocument();
            doc.Load(Assembly.GetExecutingAssembly().GetManifestResourceStream("ExcelAsyncWinForm.Ribbon.xml"));
            return doc.InnerXml;
        }

        public void HelpButton_Click(IRibbonControl control)
        {
            MessageBox.Show(control.Id);
        }

        public void UpdateButton_Click(IRibbonControl control)
        {
            MessageBox.Show(control.Id);
        }

        public void FunctionsClick(IRibbonControl control)
        {
            FrmRetrieveWeb winForm = new FrmRetrieveWeb();
            winForm.Top = 100;
            winForm.Left = 100;
            winForm.Show();
            //MessageBox.Show(control.Id);
        }

        public void FormatsClick(IRibbonControl control)
        {
            MessageBox.Show(control.Id);
        }
    }
}
