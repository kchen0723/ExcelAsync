using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading;
using System.Windows;
using System.Windows.Interop;
using System.Windows.Threading;

namespace ExcelWvvm
{
    public class WindowHelper
    {
        public delegate Window CreateWindowHandler(params object[] args);
        public delegate void SetWindowHandler(Window win, params object[] args);
        public static void ShowWindow(CreateWindowHandler createHandler, params object[] args)
        {
            ThreadStart ts = delegate { dispatchWindow(createHandler, args); };
            startThread(ts);
        }

        public static void ShowWindow<T>(SetWindowHandler setWindowHandler, params object[] args) where T : Window, new()
        {
            ThreadStart ts = delegate { dispatchWindow<T>(setWindowHandler, args); };
            startThread(ts);
        }

        private static void startThread(ThreadStart ts)
        { 
            Thread thread = new Thread(ts);
            thread.SetApartmentState(ApartmentState.STA);
            thread.IsBackground = true;
            thread.Start();
        }

        private static void dispatchWindow(CreateWindowHandler createHandler, params object[] args)
        {
            if (createHandler != null)
            {
                Window win = createHandler(args);
                win.Show();
                win.Closed += (sender, e) => win.Dispatcher.InvokeShutdown();
                Dispatcher.Run();
            }
        }

        private static void dispatchWindow<T>(SetWindowHandler setWindowHandler, params object[] args) where T : Window, new()
        {
            T win = new T();
            if (setWindowHandler != null)
            {
                setWindowHandler(win, args);
            }
            win.Show();
            win.Closed += (sender, e) => win.Dispatcher.InvokeShutdown();
            Dispatcher.Run();
        }

        public static void SetOwnerToExcel(Window win, IntPtr excelPtr)
        {
            WindowInteropHelper interop = new WindowInteropHelper(win);
            interop.Owner = excelPtr;
        }

        public static void CloseWindow(Window win)
        {
            if (win != null)
            {
                win.Dispatcher.Invoke(new Action<Window>(closeWindowByDispatcher), win);
            }
        }

        private static void closeWindowByDispatcher(Window win)
        {
            if (win != null)
            {
                win.Close();
            }
        }
    }
}
