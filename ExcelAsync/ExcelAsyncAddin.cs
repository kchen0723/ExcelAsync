using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;

using Microsoft.Office.Interop.Excel;
using ExcelDna.Integration;

namespace ExcelAsync
{
    public class ExcelAsyncAddin : IExcelAddIn
    {
        public void AutoOpen()
        {
            ExcelIntegration.RegisterUnhandledExceptionHandler(globalErrorHandler);
            AppDomain.CurrentDomain.UnhandledException += CurrentDomain_UnhandledException;
            ExcelDna.ComInterop.ComServer.DllRegisterServer();
            ExcelApp.AttachApplicationEvents();
            this.InjectWvvwDelegate();
        }

        public void AutoClose()
        {
            ExcelDna.ComInterop.ComServer.DllUnregisterServer();
        }

        private void CurrentDomain_UnhandledException(object sender, UnhandledExceptionEventArgs e)
        {
            Exception ex = e.ExceptionObject as Exception;
            System.Windows.MessageBox.Show("Excel cannot recover from error: " + ex.Message);
            Environment.Exit(-1);
        }

        public object globalErrorHandler(object ex)
        {
            return ex.ToString();
        }

        private void InjectWvvwDelegate()
        {
            ExcelWvvm.ExcelHandler.WriteToRangeHandler = ExcelManager.EntityManager.WriteToRange;
            ExcelWvvm.ExcelHandler.ShowWinHandler = WpfWindowHelper.ShowWindow;
        }
    }
}
