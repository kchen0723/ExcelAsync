using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading;
using System.Windows;
using System.Windows.Interop;
using System.Windows.Threading;

namespace ExcelAsyncWpf
{
    public class ShowWindowHelper<T> where T : Window, new()
    {
        private static Thread m_thread;

        public static void ShowWindow()
        {
            ShowWindow(true);
        }

        public static void ShowWindow(bool isSetOwnerToExcel)
        {
            m_thread = new Thread(new ParameterizedThreadStart(dipatchWindow));
            m_thread.SetApartmentState(ApartmentState.STA);
            m_thread.IsBackground = true;
            m_thread.Start(isSetOwnerToExcel);
        }

        private static void dipatchWindow(object isSetOwnerToExcel)
        {
            T win = new T();
            if (isSetOwnerToExcel.GetType() == typeof(Boolean))
            {
                bool isSetOwner = (bool)isSetOwnerToExcel;
                if (isSetOwner == true)
                {
                    WindowInteropHelper interop = new WindowInteropHelper(win);
                    interop.Owner = ExcelDna.Integration.ExcelDnaUtil.WindowHandle;
                }
            }
            win.Show();
            win.Closed += (sender, e) => win.Dispatcher.InvokeShutdown();
            Dispatcher.Run();
        }
    }
}
