using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;

using Microsoft.Office.Interop.Excel;

namespace ExcelAsyncWpf.ExcelOperate
{
    public static class ExcelApp
    {
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
    }
}
