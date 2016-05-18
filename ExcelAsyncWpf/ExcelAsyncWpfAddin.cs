using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;

using Microsoft.Office.Interop.Excel;
using ExcelDna.Integration;
using ExcelDna.ComInterop;
using ExcelAsyncWpf.ExcelOperator;

namespace ExcelAsyncWpf
{
    public class ExcelAsyncWpfAddin : IExcelAddIn
    {
        public void AutoOpen()
        {
            ExcelIntegration.RegisterUnhandledExceptionHandler(globalErrorHandler);
            ExcelApp.Application = (ExcelDnaUtil.Application as Application);
            ComServer.DllRegisterServer();
        }

        public void AutoClose()
        {
            ComServer.DllUnregisterServer();
        }

        public object globalErrorHandler(object ex)
        {
            return ex.ToString();
        }
    }
}
