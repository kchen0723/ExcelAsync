using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;

using Microsoft.Office.Interop.Excel;
using ExcelDna.Integration;
using ExcelAsyncWpf.ExcelOperate;

namespace ExcelAsyncWpf
{
    public class ExcelAsyncWpfAddin : IExcelAddIn
    {
        public void AutoOpen()
        {
            ExcelIntegration.RegisterUnhandledExceptionHandler(globalErrorHandler);
            ExcelApp.CurrentExcel = (ExcelDnaUtil.Application as Application);
        }

        public void AutoClose()
        {
            //do nothing here now
        }

        public object globalErrorHandler(object ex)
        {
            return ex.ToString();
        }
    }
}
