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
using ExcelDna.Integration.Extensibility;

namespace ExcelAsync.Ribbon
{
    [ComVisible(true)]
    public class Ribbon : ExcelRibbon
    {
        private IRibbonUI m_Ribbon;
        private bool m_IsSignalEnabled;
        public override string GetCustomUI(string RibbonID)
        {
            XmlDocument doc = new XmlDocument();
            doc.Load(Assembly.GetExecutingAssembly().GetManifestResourceStream("ExcelAsync.Ribbon.Ribbon.xml"));
            return doc.InnerXml;
        }

        public override void OnDisconnection(ext_DisconnectMode RemoveMode, ref Array custom)
        {
            MessageBox.Show("Closing Excel now");
            base.OnDisconnection(RemoveMode, ref custom);

            //Properly release com objects, see: https://www.add-in-express.com/creating-addins-blog/2013/11/05/release-excel-com-objects/
            GC.Collect();
            GC.WaitForPendingFinalizers();
            GC.Collect();
            GC.WaitForPendingFinalizers();
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
            ExcelWvvm.WindowHelper.ShowWindow<ExcelWvvm.WinRetrieveWeb>(WinRetrieveWebCreatedHandler);
        }

        public void GoogleHistoryClick(IRibbonControl control)
        {
            ExcelWvvm.WindowHelper.ShowWindow<ExcelWvvm.WinGoogleHistory>(null);
        }

        private void WinRetrieveWebCreatedHandler(Window window, params object[] args)
        {
            ExcelWvvm.WinRetrieveWeb win = window as ExcelWvvm.WinRetrieveWeb;
            if (win != null)
            {
                ExcelWvvm.WindowHelper.SetOwnerToExcel(win, ExcelDnaUtil.WindowHandle);
                win.Left = 300;
                win.Top = 300;
            }
        }

        public void FormatsClick(IRibbonControl control)
        {
            MessageBox.Show(control.Id);
        }

        public void btnClockClick(IRibbonControl control)
        {
            ExcelWvvm.WindowHelper.ShowWindow(createMvvmGoogleHistory, null);
        }

        private Window createMvvmGoogleHistory(params object[] args)
        {
            ExcelWvvm.View.WinGoogleHistory history = new ExcelWvvm.View.WinGoogleHistory();
            ExcelWvvm.ViewModel.GoogleHistoryViewModel gvm = new ExcelWvvm.ViewModel.GoogleHistoryViewModel();
            history.DataContext = gvm;
            return history;
        }

        public void RibbonUI_OnLoad(IRibbonUI ribbonUI)
        {
            m_Ribbon = ribbonUI;
            MessageBox.Show("Ribbon UI Loading");
        }

        public void btnEyeClick(IRibbonControl control)
        {
            if (this.m_Ribbon != null)
            {
                this.m_Ribbon.Invalidate();
            }
        }

        public bool OnSignalEnabled(IRibbonControl control)
        {
            m_IsSignalEnabled = !m_IsSignalEnabled;
            return m_IsSignalEnabled;
        }
    }
}
