using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading;
using System.Windows;
using System.Windows.Interop;
using System.Windows.Threading;
using ExcelWvvm;

namespace ExcelAsync
{
    public class WpfWindowHelper
    {
        private static Thread m_thread;

        public static void ShowWindow(ExcelHandler.ShowWindowHandler createdHandler, params object[] args)
        {
            ThreadStart starter = delegate { dispatchWindow(createdHandler, args); };
            m_thread = new Thread(starter);
            m_thread.SetApartmentState(ApartmentState.STA);
            m_thread.IsBackground = true;
            m_thread.Start();
        }

        public static void dispatchWindow(ExcelHandler.ShowWindowHandler handler, params object[] args)
        {
            if (handler != null)
            {
                Window win = handler(args);
                win.Show();
                win.Closed += (sender, e) => win.Dispatcher.InvokeShutdown();
                Dispatcher.Run();
            }
        }

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
            EventHandler handler = winCreatedHandler as EventHandler;
            if (handler != null)
            {
                handler(win, EventArgs.Empty);
            }
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
