using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;

using Microsoft.Office.Interop.Excel;
using Microsoft.Office.Core;

namespace ExcelAsyncWpf.ExcelOperator
{
    public static class ExcelApp
    {
        private static CommandBarButton button = null;
        public static uint ExcelMainUiThreadId { get; set; }

        private static Application m_currentExcel;
        public static Application CurrentExcel
        {
            get { return m_currentExcel; }
            set 
            {
                m_currentExcel = value;
                if (m_currentExcel != null)
                {
                    m_currentExcel.WorkbookNewSheet += new AppEvents_WorkbookNewSheetEventHandler(m_excelApp_WorkbookNewSheet);
                    m_currentExcel.WorkbookActivate += new AppEvents_WorkbookActivateEventHandler(m_excelApp_WorkbookActivate);
                }
            }
        }

        static void m_excelApp_WorkbookNewSheet(Workbook Wb, object Sh)
        {
            System.Windows.MessageBox.Show("new workbook");
        }

        private static void m_excelApp_WorkbookActivate(Workbook Wb)
        {
            System.Windows.MessageBox.Show("Workbook Activate");
        }

        public static void AddContentMenu()
        {
            CommandBar cellMenu = CurrentExcel.CommandBars["Cell"];
            //There is a bug in below line. One dot is good, too dots are bad.
            //CommandBarButton button = cellMenu.Controls.Add(Type: MsoControlType.msoControlButton, Before: cellMenu.Controls.Count, Temporary: true) as CommandBarButton;
            //button.Caption = "Test Button";
            //button.Tag = "Test Button";
            //button.OnAction = "OnButtonClick";
            button = cellMenu.Controls.Add(Type: MsoControlType.msoControlButton, Before: cellMenu.Controls.Count, Temporary: true) as CommandBarButton;
            button.Caption = "Test Button";
            button.Click += Button_Click;
        }

        private static void Button_Click(CommandBarButton Ctrl, ref bool CancelDefault)
        {
            System.Windows.MessageBox.Show("Hello from context menu");
        }

        public static void OnButtonClick()
        {
            System.Windows.MessageBox.Show("Hello from context menu");
        }
    }
}
