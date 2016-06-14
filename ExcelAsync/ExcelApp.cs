using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using Microsoft.Office.Interop.Excel;
using Microsoft.Office.Core;

namespace ExcelAsync
{
    public static class ExcelApp
    {
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
            Application.SheetBeforeRightClick += new AppEvents_SheetBeforeRightClickEventHandler(Application_SheetBeforeRightClick);
        }

        static void Application_SheetBeforeRightClick(object Sh, Range Target, ref bool Cancel)
        {
            ContextMenu.MenuManager.SetContentMenu();
        }

        static void Application_WorkbookNewSheet(Workbook Wb, object Sh)
        {
            System.Windows.MessageBox.Show("Application_WorkbookNewSheet");
        }

        private static void Application_WorkbookActivate(Workbook Wb)
        {
            ContextMenu.MenuManager.AddContextMenus();
            System.Windows.MessageBox.Show("Application_WorkbookActivate");
        }        
    }
}
