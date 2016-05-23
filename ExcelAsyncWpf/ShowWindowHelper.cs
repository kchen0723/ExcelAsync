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
    public class ShowWindowHelper
    {
        private static Thread m_thread;

        public static void ShowWindow<T>(EventHandler winCreatedHandler) where T : Window, new()
        {
            m_thread = new Thread(new ParameterizedThreadStart(dipatchWindow<T>));
            m_thread.SetApartmentState(ApartmentState.STA);
            m_thread.IsBackground = true;
            m_thread.Start(winCreatedHandler);
        }

        private static void dipatchWindow<T>(object winCreatedHandler) where T : Window, new()
        {
            T win = new T();
            (winCreatedHandler as EventHandler)?.Invoke(win, EventArgs.Empty);
            win.Show();
            win.Closed += (sender, e) => win.Dispatcher.InvokeShutdown();
            Dispatcher.Run();
        }

        public static void SetOwnerToExcel(Window win)
        {
            WindowInteropHelper interop = new WindowInteropHelper(win);
            interop.Owner = ExcelDna.Integration.ExcelDnaUtil.WindowHandle;
        }
    }
}
