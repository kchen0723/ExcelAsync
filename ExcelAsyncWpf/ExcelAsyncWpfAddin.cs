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
            AppDomain.CurrentDomain.UnhandledException += CurrentDomain_UnhandledException;
            ComServer.DllRegisterServer();
            ExcelApp.AttachApplicationEvents();
        }

        private void CurrentDomain_UnhandledException(object sender, UnhandledExceptionEventArgs e)
        {
            Exception ex = e.ExceptionObject as Exception;
            System.Windows.MessageBox.Show("Excel cannot recover from error: " + ex.Message);
            Environment.Exit(-1);
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
