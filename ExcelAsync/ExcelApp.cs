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
            //https://blogs.msdn.microsoft.com/vsofficedeveloper/2008/04/11/excel-ole-embedding-errors-if-you-have-managed-add-in-sinking-application-events-in-excel-2/
            System.Runtime.InteropServices.Marshal.ReleaseComObject(Wb);
        }

        private static void Application_WorkbookActivate(Workbook Wb)
        {
            ContextMenu.MenuManager.AddContextMenus();
            System.Windows.MessageBox.Show("Application_WorkbookActivate");
            //https://blogs.msdn.microsoft.com/vsofficedeveloper/2008/04/11/excel-ole-embedding-errors-if-you-have-managed-add-in-sinking-application-events-in-excel-2/
            System.Runtime.InteropServices.Marshal.ReleaseComObject(Wb);
        }        
    }
}
