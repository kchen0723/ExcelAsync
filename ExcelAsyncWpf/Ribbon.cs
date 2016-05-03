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
        private IRibbonUI m_Ribbon;
        private bool m_IsSignalEnabled;
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
