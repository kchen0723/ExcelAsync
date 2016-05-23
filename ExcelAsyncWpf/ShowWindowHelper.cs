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

        public static void ShowWindow(EventHandler winCreatedHandler)
        {
            m_thread = new Thread(new ParameterizedThreadStart(dipatchWindow));
            m_thread.SetApartmentState(ApartmentState.STA);
            m_thread.IsBackground = true;
            m_thread.Start(winCreatedHandler);
        }

        private static void dipatchWindow(object winCreatedHandler)
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
