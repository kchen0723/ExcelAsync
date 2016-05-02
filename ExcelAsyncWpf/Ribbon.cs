using System;
using System.Collections.Generic;
using System.Linq;
using System.Reflection;
using System.Runtime.InteropServices;
using System.Text;
using System.Xml;
using System.Windows;

using ExcelDna.Integration;
using ExcelDna.Integration.CustomUI;

namespace ExcelAsyncWpf
{
    [ComVisible(true)]
    public class Ribbon : ExcelRibbon
    {
        public override string GetCustomUI(string RibbonID)
        {
            XmlDocument doc = new XmlDocument();
            doc.Load(Assembly.GetExecutingAssembly().GetManifestResourceStream("ExcelAsyncWpf.Ribbon.xml"));
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
            ShowWindowHelper<WinRetrieveWeb>.ShowWindow();
            //MessageBox.Show(control.Id);
        }

        public void FormatsClick(IRibbonControl control)
        {
            MessageBox.Show(control.Id);
        }

        public void btnClockClick(IRibbonControl control)
        {
            ExcelOperate.ExcelApp.AddContentMenu();
        }
    }
}
