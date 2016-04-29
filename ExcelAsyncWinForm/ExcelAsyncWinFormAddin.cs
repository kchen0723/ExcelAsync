using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;

using ExcelDna.Integration;

namespace ExcelAsyncWinForm
{
    public class ExcelAsyncWinFormAddin : IExcelAddIn
    {
        public void AutoOpen()
        {
            ExcelIntegration.RegisterUnhandledExceptionHandler(globalErrorHandler);
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
