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

        public static Application Application
        {
            get
            {
                return ExcelDna.Integration.ExcelDnaUtil.Application as Application;
            }
        }

        public static void AttachApplicationEvents()
        {
            Application.WorkbookActivate += Application_WorkbookActivate;
            Application.WorkbookNewSheet += Application_WorkbookNewSheet;
        }

        static void Application_WorkbookNewSheet(Workbook Wb, object Sh)
        {
            System.Windows.MessageBox.Show("Application_WorkbookNewSheet");
        }

        private static void Application_WorkbookActivate(Workbook Wb)
        {
            System.Windows.MessageBox.Show("Application_WorkbookActivate");
        }

        public static void AddContentMenu()
        {
            CommandBar cellMenu = Application.CommandBars["Cell"];
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
